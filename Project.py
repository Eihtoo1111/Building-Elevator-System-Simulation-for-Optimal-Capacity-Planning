import simpy
import random
import math
import statistics
from collections import defaultdict, Counter
import openpyxl
import matplotlib.pyplot as plt


# BDA450 FINAL PROJECT
# Elevator Simulation using SimPy
# Floors: 0,1,2,3,4
# Building: floor 0 = garage, 1 = street/main floor, 2-4 upper floors
# Capacity per elevator = 8
# =========================================================

RANDOM_SEED = 42
SIM_TIME = 24 * 60                 # minutes in one day
FLOORS = [0, 1, 2, 3, 4]
NUM_ELEVATORS = 2                  # test with 1, 2, 3, etc.
ELEVATOR_CAPACITY = 8

# ---------- Timing assumptions (minutes) ----------
TIME_PER_FLOOR = 0.08              # about 4.8 sec per floor
DOOR_OPEN_CLOSE = 0.10             # about 6 sec
BOARD_TIME = 0.03                  # per person
ALIGHT_TIME = 0.03                 # per person
IDLE_WAIT = 0.02

# ---------- Walking assumptions ----------
HALLWAY_LENGTH_FEET = 200
AVG_WALK_DISTANCE = HALLWAY_LENGTH_FEET / 4
WALK_SPEED_FEET_PER_MIN = 250
WALK_TIME = AVG_WALK_DISTANCE / WALK_SPEED_FEET_PER_MIN



# READ EXCEL INPUT FILES
def load_counts_xlsx(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active

    rows = list(ws.iter_rows(values_only=True))
    header = rows[0]

    # expected columns:
    # [index, time, 0, 1, 2, 3, 4]
    floor_cols = {}
    for i, col in enumerate(header):
        if str(col) in {"0", "1", "2", "3", "4"} or col in [0, 1, 2, 3, 4]:
            floor_cols[int(col)] = i

    time_col = header.index("time")

    blocks = []
    for row in rows[1:]:
        if row is None:
            continue

        time_label = row[time_col]
        counts = {}
        for f in FLOORS:
            val = row[floor_cols[f]]
            counts[f] = int(val) if val is not None else 0

        blocks.append({
            "time_label": time_label,
            "counts": counts
        })

    return blocks


on_blocks = load_counts_xlsx("/Users/ACT/Downloads/Sem4/projects_to_do/SIM_project/OnCounts.xlsx")
off_blocks = load_counts_xlsx("/Users/ACT/Downloads/Sem4/projects_to_do/SIM_project/OffCounts.xlsx")

if len(on_blocks) != 96 or len(off_blocks) != 96:
    print("Warning: expected 96 time blocks, but got:")
    print("OnCounts:", len(on_blocks))
    print("OffCounts:", len(off_blocks))



# PASSENGER CLASS
class Passenger:
    def __init__(self, pid, origin, destination, arrival_time):
        self.pid = pid
        self.origin = origin
        self.destination = destination
        self.arrival_time = arrival_time
        self.request_time = None
        self.board_time = None
        self.exit_time = None
        self.direction = 1 if destination > origin else -1



# BUILD PASSENGERS FROM ON/OFF COUNTS
def choose_destination(origin, off_counts):
    """
    Choose a destination using OffCounts for the same 15-minute block.
    Same-floor trips are not allowed.
    """
    choices = []
    weights = []

    for f in FLOORS:
        if f != origin:
            choices.append(f)
            weights.append(max(off_counts.get(f, 0), 0))

    # if all weights are zero, choose any other floor uniformly
    if sum(weights) == 0:
        return random.choice(choices)

    return random.choices(choices, weights=weights, k=1)[0]


def build_daily_passenger_schedule(on_blocks, off_blocks):
    """
    Creates a list of Passenger objects from Excel traffic inputs.

    Since the inputs only provide:
    - counts of people getting ON by floor
    - counts of people getting OFF by floor

    exact origin-destination pairs are not known.
    So destinations are assigned probabilistically using OffCounts
    within the same 15-minute block.
    """
    passengers = []
    pid = 0

    for block_idx in range(min(len(on_blocks), len(off_blocks))):
        block_start = block_idx * 15
        block_end = block_start + 15

        on_counts = on_blocks[block_idx]["counts"]
        off_counts = off_blocks[block_idx]["counts"]

        for origin in FLOORS:
            num_boarding = on_counts.get(origin, 0)

            for _ in range(num_boarding):
                arrival_time = random.uniform(block_start, block_end)
                destination = choose_destination(origin, off_counts)

                passengers.append(
                    Passenger(
                        pid=pid,
                        origin=origin,
                        destination=destination,
                        arrival_time=arrival_time
                    )
                )
                pid += 1

    passengers.sort(key=lambda p: p.arrival_time)
    return passengers



# ELEVATOR SYSTEM
class ElevatorSystem:
    def __init__(self, env, num_elevators):
        self.env = env
        self.num_elevators = num_elevators

        self.waiting_up = {f: [] for f in FLOORS}
        self.waiting_down = {f: [] for f in FLOORS}

        self.elevators = []
        self.stats = []

        for eid in range(num_elevators):
            elev = Elevator(env, eid, self)
            self.elevators.append(elev)
            env.process(elev.run())

    def add_passenger(self, passenger):
        passenger.request_time = self.env.now + WALK_TIME
        if passenger.direction == 1:
            self.waiting_up[passenger.origin].append(passenger)
        else:
            self.waiting_down[passenger.origin].append(passenger)

    def has_waiting(self):
        for f in FLOORS:
            if self.waiting_up[f] or self.waiting_down[f]:
                return True
        return False

    def get_waiting_count(self, floor, direction):
        if direction == 1:
            return len(self.waiting_up[floor])
        return len(self.waiting_down[floor])

    def board_passengers(self, floor, direction, available_space):
        if direction == 1:
            queue = self.waiting_up[floor]
        else:
            queue = self.waiting_down[floor]

        boarded = []
        while queue and len(boarded) < available_space:
            boarded.append(queue.pop(0))

        return boarded

    def record_passenger(self, p):
        wait_time = p.board_time - p.request_time
        ride_time = p.exit_time - p.board_time
        total_time = p.exit_time - p.arrival_time

        self.stats.append({
            "pid": p.pid,
            "origin": p.origin,
            "destination": p.destination,
            "arrival_time": p.arrival_time,
            "request_time": p.request_time,
            "board_time": p.board_time,
            "exit_time": p.exit_time,
            "wait_time": wait_time,
            "ride_time": ride_time,
            "total_time": total_time,
            "time_block": int(p.arrival_time // 15)
        })


class Elevator:
    def __init__(self, env, eid, system):
        self.env = env
        self.eid = eid
        self.system = system
        self.current_floor = 0
        self.direction = 1
        self.passengers = []
        self.capacity = ELEVATOR_CAPACITY

    def move_one_floor(self, target_floor):
        travel = abs(target_floor - self.current_floor) * TIME_PER_FLOOR
        if travel > 0:
            yield self.env.timeout(travel)
            self.current_floor = target_floor

    def unload_passengers(self):
        leaving = [p for p in self.passengers if p.destination == self.current_floor]
        if leaving:
            yield self.env.timeout(DOOR_OPEN_CLOSE)
            for p in leaving:
                yield self.env.timeout(ALIGHT_TIME)
                p.exit_time = self.env.now
                self.system.record_passenger(p)
                self.passengers.remove(p)

    def load_passengers(self):
        waiting_here = self.system.get_waiting_count(self.current_floor, self.direction)
        if waiting_here > 0 and len(self.passengers) < self.capacity:
            boarded = self.system.board_passengers(
                self.current_floor,
                self.direction,
                self.capacity - len(self.passengers)
            )
            if boarded:
                yield self.env.timeout(DOOR_OPEN_CLOSE)
                for p in boarded:
                    yield self.env.timeout(BOARD_TIME)
                    p.board_time = self.env.now
                    self.passengers.append(p)

    def should_stop_here(self):
        dropoff = any(p.destination == self.current_floor for p in self.passengers)
        pickup = (
            self.system.get_waiting_count(self.current_floor, self.direction) > 0
            and len(self.passengers) < self.capacity
        )
        return dropoff or pickup

    def requests_ahead(self):
        if self.direction == 1:
            for f in range(self.current_floor + 1, max(FLOORS) + 1):
                if any(p.destination == f for p in self.passengers):
                    return True
                if self.system.get_waiting_count(f, 1) > 0:
                    return True
        else:
            for f in range(self.current_floor - 1, min(FLOORS) - 1, -1):
                if any(p.destination == f for p in self.passengers):
                    return True
                if self.system.get_waiting_count(f, -1) > 0:
                    return True
        return False

    def opposite_requests_exist(self):
        if self.direction == 1:
            for f in range(self.current_floor, min(FLOORS) - 1, -1):
                if self.system.get_waiting_count(f, -1) > 0:
                    return True
        else:
            for f in range(self.current_floor, max(FLOORS) + 1):
                if self.system.get_waiting_count(f, 1) > 0:
                    return True
        return False

    def nearest_request_floor(self):
        candidate_floors = []
        for f in FLOORS:
            if self.system.waiting_up[f] or self.system.waiting_down[f]:
                candidate_floors.append(f)

        if not candidate_floors:
            return 0

        return min(candidate_floors, key=lambda f: abs(f - self.current_floor))

    def run(self):
        while True:
            if self.should_stop_here():
                yield from self.unload_passengers()
                yield from self.load_passengers()

            if self.passengers or self.system.has_waiting():
                if self.requests_ahead():
                    next_floor = self.current_floor + self.direction
                    if next_floor in FLOORS:
                        yield from self.move_one_floor(next_floor)
                    else:
                        self.direction *= -1
                elif self.opposite_requests_exist():
                    self.direction *= -1
                else:
                    target = self.nearest_request_floor()
                    if target != self.current_floor:
                        step = 1 if target > self.current_floor else -1
                        self.direction = step
                        yield from self.move_one_floor(self.current_floor + step)
                    else:
                        yield self.env.timeout(IDLE_WAIT)
            else:
                # idle return to floor 0
                if self.current_floor != 0:
                    step = -1 if self.current_floor > 0 else 1
                    self.direction = step
                    yield from self.move_one_floor(self.current_floor + step)
                else:
                    yield self.env.timeout(IDLE_WAIT)



# PASSENGER ARRIVAL PROCESS
def passenger_arrivals(env, system, passengers):
    for p in passengers:
        yield env.timeout(max(0, p.arrival_time - env.now))
        system.add_passenger(p)



# RESULTS
def percentile(data, p):
    if not data:
        return 0
    data = sorted(data)
    k = (len(data) - 1) * p / 100
    f = math.floor(k)
    c = math.ceil(k)
    if f == c:
        return data[int(k)]
    return data[f] * (c - k) + data[c] * (k - f)


def summarize_results(stats, num_elevators):
    if not stats:
        print("No passengers served.")
        return

    wait_times = [x["wait_time"] for x in stats]
    ride_times = [x["ride_time"] for x in stats]
    total_times = [x["total_time"] for x in stats]

    print("\n" + "=" * 50)
    print(f"RESULTS FOR {num_elevators} ELEVATOR(S)")
    print("=" * 50)
    print(f"Passengers served: {len(stats)}")
    print(f"Average wait time: {statistics.mean(wait_times):.2f} min")
    print(f"Average ride time: {statistics.mean(ride_times):.2f} min")
    print(f"Average total time: {statistics.mean(total_times):.2f} min")
    print(f"Std dev wait time: {statistics.pstdev(wait_times):.2f} min")
    print(f"90th percentile wait: {percentile(wait_times, 90):.2f} min")
    print(f"95th percentile wait: {percentile(wait_times, 95):.2f} min")
    print(f"Max wait time: {max(wait_times):.2f} min")

    # Histograms
    plt.hist(wait_times, bins=50)
    plt.title(f"Wait Time/ Total Time Distribution ({num_elevators} Elevators)")
    plt.xlabel("Wait Time (minutes)")
    plt.ylabel("Frequency")
    plt.show()

    by_floor = defaultdict(list)
    for row in stats:
        by_floor[row["origin"]].append(row["wait_time"])

    print("\nWait time by origin floor:")
    for f in sorted(by_floor):
        vals = by_floor[f]
        print(
            f"Floor {f}: avg={statistics.mean(vals):.2f}, "
            f"p90={percentile(vals, 90):.2f}, "
            f"max={max(vals):.2f}, n={len(vals)}"
        )

    by_block = defaultdict(list)
    for row in stats:
        by_block[row["time_block"]].append(row["wait_time"])

    print("\nWait time by 15-minute block:")
    for b in sorted(by_block):
        vals = by_block[b]
        hh = (b * 15) // 60
        mm = (b * 15) % 60
        print(f"{hh:02d}:{mm:02d}  avg={statistics.mean(vals):.2f}  n={len(vals)}")



# MAIN SIMULATION RUNNER
def run_simulation(num_elevators, seed=42):
    random.seed(seed)

    daily_passengers = build_daily_passenger_schedule(on_blocks, off_blocks)

    env = simpy.Environment()
    system = ElevatorSystem(env, num_elevators)

    env.process(passenger_arrivals(env, system, daily_passengers))
    env.run(until=SIM_TIME + 180)

    summarize_results(system.stats, num_elevators)
    return system.stats



# RUN EXPERIMENTS
if __name__ == "__main__":
    stats_1 = run_simulation(num_elevators=1, seed=RANDOM_SEED)
    stats_2 = run_simulation(num_elevators=2, seed=RANDOM_SEED)
    stats_3 = run_simulation(num_elevators=3, seed=RANDOM_SEED)