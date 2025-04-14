import sys
import subprocess

try:
    import ortools

    print("OR-Tools is already installed.")
except ImportError:
    print("OR-Tools is not installed. Installing now...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "ortools"])
    import ortools

    print("OR-Tools has been successfully installed.")

import os
import pandas as pd
import numpy as np
import math
from ortools.constraint_solver import pywrapcp, routing_enums_pb2

"""STEP 1: LOAD EXCEL DATA"""

file_path = ("Schneider_Electric_VRP_Model.xlsx")  # Update with your actual file path)

# Load sheets using openpyxl to preserve formulas
xls = pd.ExcelFile(file_path, engine="openpyxl")

# Read customer name and index
customer_df = pd.read_excel(xls, "Distance (KM)", engine="openpyxl")
customer_id_to_name_map = dict(zip(customer_df["S/N"].values, customer_df["Customer name"].values))

# Read distance matrix
distance_df = pd.read_excel(xls, "Distance (KM)", engine="openpyxl")
# Assumes the first row/column are headers; remove them
# Note that the index here is the customer ID
distance_matrix = distance_df.iloc[1:, 2:].values

# Read demand (assumed to include depot at index 0)
demand_df = pd.read_excel(xls, "Demand (KG)", engine="openpyxl")

# Note that the index of the list is the drop-off ID, NOT the customer ID
demand = demand_df["Demand (KG)"].values
demand_with_customer_id = list(zip(demand_df['S/N'], demand_df['Demand (KG)']))

# Track invalid customer with demand 0 except depot
invalid_nodes = [index for index, demand in enumerate(demand) if index != 0 and demand == 0]

# Read truck capacities and fixed costs
trucks_df = pd.read_excel(xls, "Trucks", index_col=0, engine="openpyxl", dtype={"Amount": object})
truck_capacity = trucks_df["Capacity (KG)"].values
truck_fixed_cost = trucks_df["Fixed Cost (KRW)"].values

# Read per km rate from Parameters sheet
parameters_df = pd.read_excel(xls, "Parameters", index_col=0, engine="openpyxl")
per_km_rate = int(parameters_df.iloc[0, 0])  # e.g., value from cell B2 (KRW/km) , convert to int explicitly
speed_fast = int(parameters_df.iloc[1, 0])  # e.g., value from cell B3 (km/h), convert to int explicitly
speed_slow = int(parameters_df.iloc[2, 0])  # e.g., value from cell B3 (km/h), convert to int explicitly
base_unloading_time = int(parameters_df.iloc[3, 0])  # e.g., value from cell B5 (min), convert to int explicitly
unloading_time_scale_factor = int(parameters_df.iloc[4, 0])  # e.g., value from cell B6 (min), convert to int explicitly
lunch_break = int(parameters_df.iloc[5, 0])  # e.g., value from cell B7 (min), convert to int explicitly
waiting_time = int(parameters_df.iloc[6, 0])  # e.g., value from cell B8 (min), convert to int explicitly

#Calculate total unloading time based on demand in kg.
def get_total_unloading_time(demand_kg):
    if demand_kg == 0:
        return 0
    return base_unloading_time + (unloading_time_scale_factor * math.ceil(math.log(demand_kg + 1)))

# === Read only the two columns for time windows from the Time Constraint sheet ===
time_df = pd.read_excel(
    xls,
    "Time Constraint",
    usecols=["S/N", "Start Time Window (HH:MM)", "End Time Window (HH:MM)"],
    engine="openpyxl",)

# Helper function to convert a time value (HH:MM) or datetime to minutes
def convert_time_to_minutes(t):
    if isinstance(t, str):
        h, m = t.split(":")
        return int(h) * 60 + int(m)
    elif hasattr(t, "strftime"):
        s = t.strftime("%H:%M")
        h, m = s.split(":")
        return int(h) * 60 + int(m)
    elif hasattr(t, "hour"):
        return t.hour * 60 + t.minute
    elif np.issubdtype(t.dtype, np.datetime64):
        timestamp = pd.Timestamp(t)
        return timestamp.hour * 60 + timestamp.minute
    else:
        return int(t)

# Convert the start/end time columns to integer minutes
start_time_minutes = [convert_time_to_minutes(x) for x in time_df["Start Time Window (HH:MM)"].values]
end_time_minutes = [convert_time_to_minutes(x) for x in time_df["End Time Window (HH:MM)"].values]

# Utility functions to get customer ID from drop-off node and vice versa
def get_customer_id_from_demand_node(node):
    customer_id = demand_with_customer_id[node][0]
    return customer_id

def get_primary_customer_id(customer_id):
    customer_id_string = str(customer_id)
    if '.' in customer_id_string:
        primary_customer_id = int(customer_id_string.split('.')[0])
    else:
        primary_customer_id = customer_id
    return primary_customer_id

# customer_id can be either the primary customer ID (10) or the sub-customer ID (10.1, 10.2, etc.)
def get_time_window_for_customer_by_id(customer_id):
    primary_customer_id = get_primary_customer_id(customer_id)

    # Get the row index for the customer ID
    row = time_df.loc[time_df['S/N'] == customer_id]

    # If there is no entry for the customer ID, fall back to the primary customer ID
    if row.empty:
        row = time_df.loc[time_df['S/N'] == primary_customer_id]

        if row.empty:
            print(f"No entry found for customer_id: {customer_id}")
            return None  # or handle the case where the customer_id is not found

    row_index = row.index.values[0]
    return start_time_minutes[row_index], end_time_minutes[row_index]

def get_time_window_for_customer_by_node(node):
    customer_id = get_customer_id_from_demand_node(node)
    return get_time_window_for_customer_by_id(customer_id)

def get_distance_between_customers_by_id(customer_id_1, customer_id_2):
    # Get the primary customer IDs
    primary_customer_id_1 = get_primary_customer_id(customer_id_1)
    primary_customer_id_2 = get_primary_customer_id(customer_id_2)

    # distance_df has a column 'KM' to store customer IDs and the first row as headers
    # 'S/N' is used to identify customer IDs for time windows, let's use that consistently

    # Ensure customer_id_1 exists in the DataFrame, using 'S/N' column instead of 'KM'
    if primary_customer_id_1 not in distance_df['S/N'].values:
        raise ValueError(f"Customer ID {customer_id_1} not found in distance_df.")

    # Ensure customer_id_2 exists as a column in the DataFrame
    if primary_customer_id_2 not in distance_df.columns:
        raise ValueError(f"Column {customer_id_2} not found in distance_df.")

    # Get the distance value using .loc with 'S/N'
    distance_value = distance_df.loc[distance_df['S/N'] == primary_customer_id_1, primary_customer_id_2].values[0]
    return distance_value

def get_distance_between_customers_by_node(node_1, node_2):
    customer_id_1 = get_customer_id_from_demand_node(node_1)
    customer_id_2 = get_customer_id_from_demand_node(node_2)
    return get_distance_between_customers_by_id(customer_id_1, customer_id_2)

num_trucks = len(truck_capacity)
num_drop_offs = len(demand_with_customer_id) # Includes depot at index 0
depot = 0  # Depot is at index 0

"""STEP 2: SOLVE THE VRP WITH STRICT TIME WINDOW CONSTRAINTS"""


def solve_vrp(
    demand,
    truck_capacity,
    truck_fixed_cost,
    per_km_rate,
    num_trucks,
    speed_fast,
    speed_slow,
):
    manager = pywrapcp.RoutingIndexManager(num_drop_offs, num_trucks, depot)
    routing = pywrapcp.RoutingModel(manager)

    # Demand callback and capacity constraint
    def demand_callback(from_index):
        from_node = manager.IndexToNode(int(from_index))
        return demand[from_node]

    demand_callback_index = routing.RegisterUnaryTransitCallback(demand_callback)
    routing.AddDimensionWithVehicleCapacity(demand_callback_index, 0, truck_capacity, True, "Capacity")

    # Time callback: computes travel + unloading time in minutes
    def time_callback(from_index, to_index):
        if to_index in invalid_nodes:
            return 0

        from_node = manager.IndexToNode(int(from_index))
        to_node = manager.IndexToNode(int(to_index))
        dist = get_distance_between_customers_by_node(from_node, to_node)

        def get_travel_time(distance):
            if distance > 70:
                return (distance / speed_fast) * 60
            else:
                return (distance / speed_slow) * 60

        travel_time = get_travel_time(dist)
        unloading_time = get_total_unloading_time(demand[to_node])
        return math.ceil(travel_time + unloading_time)

    time_callback_index = routing.RegisterTransitCallback(time_callback)
    horizon = 1440  # 24-hour horizon in minutes
    routing.AddDimension(
        time_callback_index,
        waiting_time,  # slack (waiting time)(can be modified by user)
        horizon,  # maximum time per vehicle
        False,  # do not force start cumul to zero
        "Time",
    )
    time_dimension = routing.GetDimensionOrDie("Time")

    # Enforce strict time windows:
    # First set the depot node itself
    index_depot = manager.NodeToIndex(depot)
    time_dimension.CumulVar(index_depot).SetRange(540, horizon)

    # Set ALL vehicle start nodes to 540
    for vehicle_id in range(num_trucks):
        index_start = routing.Start(vehicle_id)
        time_dimension.CumulVar(index_start).SetRange(540, horizon)
    # Set customer time windows
    for i in range(1, num_drop_offs):  # Skip depot (0)
        index_i = manager.NodeToIndex(i)
        start_tw, end_tw = get_time_window_for_customer_by_node(i) # e.g., 12:00 in minutes
        unloading = get_total_unloading_time(demand[i])

        adjusted_start_tw = start_tw + unloading
        # Check if lunch break is possible by checking the duration between start time (incl. unloading) and end time
        is_break_allow = end_tw - adjusted_start_tw >= lunch_break
        # If lunch break is possible, we allow time for possible lunch break by reducing the end time (to avoid exceed end time)
        adjusted_end_tw = end_tw - lunch_break if is_break_allow else end_tw
        time_dimension.CumulVar(index_i).SetRange(adjusted_start_tw, adjusted_end_tw)

        # Encourage the solver to minimize the departure time without forcing it exactly.
        routing.AddVariableMinimizedByFinalizer(time_dimension.CumulVar(index_i))

    # This prioritises the time as the main cost constraint
    time_callback_index = routing.RegisterTransitCallback(time_callback)
    routing.SetArcCostEvaluatorOfAllVehicles(time_callback_index)

    # Adjust search parameters to encourage reallocation:
    search_parameters = pywrapcp.DefaultRoutingSearchParameters()
    search_parameters.first_solution_strategy = (routing_enums_pb2.FirstSolutionStrategy.SAVINGS)
    search_parameters.local_search_metaheuristic = (routing_enums_pb2.LocalSearchMetaheuristic.AUTOMATIC)
    search_parameters.time_limit.seconds = 60
    search_parameters.use_full_propagation = True

    solution = routing.SolveWithParameters(search_parameters)
    if solution:
        return extract_solution(
            manager,
            routing,
            solution,
            truck_fixed_cost,
            demand,
            per_km_rate,
            time_dimension
        )
    else:
        return None


"""STEP 3: EXTRACT SOLUTION, COMPUTE CUMULATIVE TIME, FORMAT ARRIVAL TIME"""


def extract_solution(
    manager,
    routing,
    solution,
    truck_fixed_cost,
    demand,
    per_km_rate,
    time_dimension
):
    routes = {}
    cumulative_rows = []
    total_cost = 0.0
    total_distance = 0.0
    nodes_visited_count = {}

    for vehicle_id in range(manager.GetNumberOfVehicles()):
        index = routing.Start(vehicle_id)
        route = []
        total_demand = 0.0
        route_distance = 0.0
        variable_cost = 0.0
        prev_node = None
        stop_number = 0

        break_taken = False  # put this before the while loop

        while not routing.IsEnd(index):
            node = manager.IndexToNode(int(index))
            if node in invalid_nodes:
                index = solution.Value(routing.NextVar(index))
                continue
            is_break_node = "N"
            customer_id = get_customer_id_from_demand_node(node)

            # Format customer_id to clearly indicate when a customer is visited on multiple drop-offs
            if node != depot:
                if customer_id in nodes_visited_count:
                    original_customer_id = customer_id
                    customer_id = f'{original_customer_id}.{nodes_visited_count[original_customer_id]}'
                    nodes_visited_count[original_customer_id] += 1
                else:
                    nodes_visited_count[customer_id] = 1
            route.append(customer_id)
            current_demand = demand[node]
            total_demand += current_demand

            # Get the actual time from the solver's dimension
            time_var = time_dimension.CumulVar(index)

            if break_taken:
                total_time = solution.Min(time_var) + lunch_break
            else:
                total_time = solution.Value(time_var)

            # Calculate arrival time by subtracting unloading time
            unloading_time = get_total_unloading_time(current_demand)
            arrival_time = total_time - unloading_time
            depot_time = solution.Value(
                time_dimension.CumulVar(routing.Start(vehicle_id))
            )
            departure_time = total_time

            if stop_number == 0:
                cum_time_hr = 0.0
            else:
                cum_time_hr = round(departure_time / 60.0, 2) - round(
                    depot_time / 60.0, 2
                )

            if not break_taken and cum_time_hr >= 4:
                unloading_time += lunch_break
                departure_time += lunch_break
                break_taken = True  # only allow it once
                is_break_node = "Y"

            hh_dep = departure_time // 60
            mm_dep = departure_time % 60
            departure_str = f"{hh_dep:02d}:{mm_dep:02d}"

            # Convert to hours and format string
            hh_arrival = arrival_time // 60
            mm_arrival = arrival_time % 60
            arrival_str = f"{hh_arrival:02d}:{mm_arrival:02d}"
            hh_departure = departure_time // 60
            mm_departure = departure_time % 60
            departure_str = f"{hh_departure:02d}:{mm_departure:02d}"

            seg_distance = 0.0
            seg_cost = 0.0
            if prev_node is not None:
                seg_distance = get_distance_between_customers_by_node(prev_node, node)
                seg_cost = seg_distance * per_km_rate
                route_distance += seg_distance
                variable_cost += seg_cost

            cumulative_rows.append(
                (
                    vehicle_id + 1,
                    stop_number,
                    customer_id,
                    current_demand,
                    seg_distance,
                    seg_cost,
                    cum_time_hr,
                    arrival_str,
                    departure_str,
                    unloading_time,
                    is_break_node,
                )
            )

            stop_number += 1
            prev_node = node
            index = solution.Value(routing.NextVar(index))

        if len(route) > 1:
            fixed_cost = truck_fixed_cost[vehicle_id]
            total_cost += variable_cost + fixed_cost
            total_distance += route_distance
            routes[vehicle_id] = (route, total_demand, variable_cost, fixed_cost)

    print(f"\nTotal Distance for all Trucks: {total_distance:.2f} km")
    print(f"Total Cost (Variable + Fixed): KRW{total_cost:.2f}")

    cumulative_df = pd.DataFrame(
        cumulative_rows,
        columns=[
            "Truck ID",
            "Stop Order",
            "Node",
            "Demand (KG)",
            "Segment Distance (KM)",
            "Segment Cost (KRW)",
            "Computed Cumulative Time (hr)",
            "Arrival Time (HH:MM)",
            "Departure Time (HH:MM)",
            "Unloading Time (min)",
            "Lunch",
        ],
    )

    return routes, cumulative_df


"""STEP 4: RUN THE SOLVER AND SAVE RESULTS"""

result = solve_vrp(
    demand,
    truck_capacity,
    truck_fixed_cost,
    per_km_rate,
    num_trucks,
    speed_fast,
    speed_slow,
)

if result:
    routes, cumulative_df = result

    # Build summary DataFrame (VRP Results)
    routes_df = pd.DataFrame(
        [
            (
                truck + 1,
                " -> ".join(map(str, route)),
                capacity_used,
                variable_cost,
                fixed_cost,
                variable_cost + fixed_cost,
            )
            for truck, (
                route,
                capacity_used,
                variable_cost,
                fixed_cost,
            ) in routes.items()
        ],
        columns=[
            "Truck",
            "Route",
            "Capacity Used (KG)",
            "Variable Cost (KRW)",
            "Fixed Cost (KRW)",
            "Total Cost (KRW)",
        ],
    )

    total_capacity_all_trucks = routes_df["Capacity Used (KG)"].sum()
    total_variable_cost_all_trucks = routes_df["Variable Cost (KRW)"].sum()
    total_fixed_cost_all_trucks = routes_df["Fixed Cost (KRW)"].sum()
    total_cost_all_trucks = routes_df["Total Cost (KRW)"].sum()

    # Add total cost row to routes_df
    total_row = pd.DataFrame(
        [
            {
                "Truck": "Total",
                "Capacity Used (KG)": total_capacity_all_trucks,
                "Variable Cost (KRW)": total_variable_cost_all_trucks,
                "Fixed Cost (KRW)": total_fixed_cost_all_trucks,
                "Total Cost (KRW)": total_cost_all_trucks,
            }
        ]
    )
    routes_df = pd.concat([routes_df, total_row], ignore_index=True)

    os.system(
        f"cp '{file_path}' 'Solved_VRP.xlsx'"
    )  # Create a copy of the original file

    # Write both sheets to Excel
    with pd.ExcelWriter(
        "Solved_VRP.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace"
    ) as writer:
        routes_df.to_excel(writer, sheet_name="VRP Results", index=False)
        cumulative_df.to_excel(
            writer, sheet_name="Cumulative Delivery Times", index=False
        )
    from openpyxl import load_workbook

    # Load the written workbook
    wb = load_workbook("Solved_VRP.xlsx")

    for sheet_name in ["VRP Results", "Cumulative Delivery Times"]:
        ws = wb[sheet_name]
        for column_cells in ws.columns:
            max_length = 0
            column = column_cells[0].column_letter  # Get the column letter (e.g., 'A')
            for cell in column_cells:
                try:
                    cell_value = str(cell.value)
                    if len(cell_value) > max_length:
                        max_length = len(cell_value)
                except:
                    pass
            adjusted_width = max_length + 2
            ws.column_dimensions[column].width = adjusted_width

    # Save with adjusted widths
    wb.save("Solved_VRP.xlsx")

    print("✅ Results saved to Solved_VRP.xlsx!")
else:
    print("❌ No solution found.")