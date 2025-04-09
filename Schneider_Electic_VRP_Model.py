# -*- coding: utf-8 -*-
"""Jay Edit

Automatically generated by Colab.

Original file is located at
    https://colab.research.google.com/drive/12vfmwgga2Z1n-5gPqFoxQkl_TyDh8Bbw
"""

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

file_path = (
    "Modelling_Using_Python1_clean - Copy (1).xlsx"  # Update with your actual file path
)

# Load sheets using openpyxl to preserve formulas
xls = pd.ExcelFile(file_path, engine="openpyxl")

# Read distance matrix
distance_df = pd.read_excel(xls, "Distance (KM)", index_col=0, engine="openpyxl")
# Assumes the first row/column are headers; remove them
distance_matrix = distance_df.iloc[1:, 1:].values

# Read demand (assumed to include depot at index 0)
demand_df = pd.read_excel(xls, "Demand (KG)", index_col=0, engine="openpyxl")
demand = demand_df["Demand (KG)"].values

# Read truck capacities and fixed costs
trucks_df = pd.read_excel(xls, "Trucks", index_col=0, engine="openpyxl", dtype={'Amount': object})
truck_capacity = trucks_df["Capacity (KG)"].values
truck_fixed_cost = trucks_df["Fixed Cost ($)"].values

# Read per km rate from Parameters sheet
parameters_df = pd.read_excel(xls, "Parameters", index_col=0, engine="openpyxl")

per_km_rate = int(parameters_df.iloc[0, 0])  # e.g., value from cell B2 ($/km) , convert to int explicitly
speed_fast = int(parameters_df.iloc[1, 0])  # e.g., value from cell B3 (km/h), convert to int explicitly
speed_slow = int(parameters_df.iloc[2, 0])  # e.g., value from cell B3 (km/h), convert to int explicitly
base_unloading_time = int(parameters_df.iloc[3, 0])  # e.g., value from cell B5 (min), convert to int explicitly
unloading_time_scale_factor = int(parameters_df.iloc[4, 0])  # e.g., value from cell B6 (min), convert to int explicitly
lunch_break = int(parameters_df.iloc[5, 0])  # e.g., value from cell B7 (min), convert to int explicitly
waiting_time = int(parameters_df.iloc[6, 0])  # e.g., value from cell B8 (min), convert to int explicitly

"""
Calculate total unloading time based on demand in kg.

:param demand_kg: Weight in kilograms
:return: Total nploading time in minutes
"""

def get_total_unloading_time(demand_kg):
    if demand_kg == 0:
        return 0
    return base_unloading_time + (
        unloading_time_scale_factor * math.ceil(math.log(demand_kg + 1))
    )


# === Read only the two columns for time windows from the Time Constraint sheet ===
time_df = pd.read_excel(
    xls,
    "Time Constraint",
    usecols=["Start Time Window (HH:MM)", "End Time Window (HH:MM)"],
    engine="openpyxl",
)


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
start_time_minutes = [
    convert_time_to_minutes(x) for x in time_df["Start Time Window (HH:MM)"].values
]
end_time_minutes = [
    convert_time_to_minutes(x) - lunch_break
    for x in time_df["End Time Window (HH:MM)"].values
]

num_trucks = len(truck_capacity)
num_locations = len(distance_matrix)  # Includes depot at index 0
depot = 0  # Depot is at index 0

print("Truck ID, Capacities, and Fixed Costs:")
for i, (cap, cost) in enumerate(zip(truck_capacity, truck_fixed_cost)):
    print(f"Truck {i+1}: {cap} KG, Fixed Cost: ${cost:.2f}")

"""STEP 2: SOLVE THE VRP WITH STRICT TIME WINDOW CONSTRAINTS"""

def solve_vrp(
    distance_matrix,
    demand,
    truck_capacity,
    truck_fixed_cost,
    per_km_rate,
    num_trucks,
    start_time_minutes,
    end_time_minutes,
    speed_fast,
    speed_slow,
):
    manager = pywrapcp.RoutingIndexManager(num_locations, num_trucks, depot)
    routing = pywrapcp.RoutingModel(manager)

    # Distance callback and cost
    def distance_callback(from_index, to_index):
        from_node = manager.IndexToNode(int(from_index))
        to_node = manager.IndexToNode(int(to_index))
        return distance_matrix[from_node][to_node]

    # This prioritises the distance as the main cost constraint
    # transit_callback_index = routing.RegisterTransitCallback(distance_callback)
    # routing.SetArcCostEvaluatorOfAllVehicles(transit_callback_index)

    # Demand callback and capacity constraint
    def demand_callback(from_index):
        from_node = manager.IndexToNode(int(from_index))
        return demand[from_node]

    demand_callback_index = routing.RegisterUnaryTransitCallback(demand_callback)
    routing.AddDimensionWithVehicleCapacity(
        demand_callback_index, 0, truck_capacity, True, "Capacity"
    )

    # Time callback: computes travel + unloading time in minutes
    def time_callback(from_index, to_index):
        from_node = manager.IndexToNode(int(from_index))
        to_node = manager.IndexToNode(int(to_index))
        dist = distance_matrix[from_node][to_node]

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
    for i in range(1, num_locations):  # Skip depot (0)
        index_i = manager.NodeToIndex(i)
        start_tw = start_time_minutes[i - 1]  # e.g., 12:00 in minutes
        end_tw = end_time_minutes[i - 1]  # e.g., 14:00 in minutes
        unloading = get_total_unloading_time(demand[i])

        # Force the departure time to be no earlier than (start_tw + unloading),
        # so that arrival time (departure - unloading) will be ≥ start_tw.
        time_dimension.CumulVar(index_i).SetRange(start_tw + unloading, end_tw)

        # Encourage the solver to minimize the departure time without forcing it exactly.
        routing.AddVariableMinimizedByFinalizer(time_dimension.CumulVar(index_i))

    # This prioritises the time as the main cost constraint
    time_callback_index = routing.RegisterTransitCallback(time_callback)
    routing.SetArcCostEvaluatorOfAllVehicles(time_callback_index)

    # Adjust search parameters to encourage reallocation:
    search_parameters = pywrapcp.DefaultRoutingSearchParameters()
    search_parameters.first_solution_strategy = (
        routing_enums_pb2.FirstSolutionStrategy.SAVINGS
    )
    search_parameters.local_search_metaheuristic = (
        routing_enums_pb2.LocalSearchMetaheuristic.AUTOMATIC
    )
    search_parameters.time_limit.seconds = 60
    search_parameters.solution_limit = 1000
    search_parameters.use_full_propagation = True

    solution = routing.SolveWithParameters(search_parameters)
    if solution:
        return extract_solution(
            manager,
            routing,
            solution,
            truck_fixed_cost,
            demand,
            distance_matrix,
            per_km_rate,
            time_dimension,
            speed_fast,
            speed_slow,
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
    distance_matrix,
    per_km_rate,
    time_dimension,
    speed_fast,
    speed_slow,
):
    routes = {}
    cumulative_rows = []
    total_cost = 0.0
    total_distance = 0.0

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
            is_break_node = "N"
            node = manager.IndexToNode(int(index))
            route.append(node)
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
                cum_time_hr = round(arrival_time / 60.0, 2) - round(
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
                seg_distance = distance_matrix[prev_node][node]
                seg_cost = seg_distance * per_km_rate
                route_distance += seg_distance
                variable_cost += seg_cost

            cumulative_rows.append(
                (
                    vehicle_id + 1,
                    stop_number,
                    node,
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
    print(f"Total Cost (Variable + Fixed): ${total_cost:.2f}")

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
    distance_matrix,
    demand,
    truck_capacity,
    truck_fixed_cost,
    per_km_rate,
    num_trucks,
    start_time_minutes,
    end_time_minutes,
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
            "Variable Cost ($)",
            "Fixed Cost ($)",
            "Total Cost ($)",
        ],
    )

    total_capacity_all_trucks = routes_df["Capacity Used (KG)"].sum()
    total_variable_cost_all_trucks = routes_df["Variable Cost ($)"].sum()
    total_fixed_cost_all_trucks = routes_df["Fixed Cost ($)"].sum()
    total_cost_all_trucks = routes_df["Total Cost ($)"].sum()

    # Add total cost row to routes_df
    total_row = pd.DataFrame([{
        "Truck": "Total",
        "Capacity Used (KG)": total_capacity_all_trucks,
        "Variable Cost ($)": total_variable_cost_all_trucks,
        "Fixed Cost ($)": total_fixed_cost_all_trucks,
        "Total Cost ($)": total_cost_all_trucks
    }])
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

    print("✅ Results saved to Solved_VRP.xlsx!")
else:
    print("❌ No solution found.")