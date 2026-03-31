import simpy
import random

rand = random.Random()

class ElevatorEventSim:

    def runsim(self, num_people, mean_arrival, service_time):
        env = simpy.Environment()
        env.process(self.source(env, num_people, mean_arrival, service_time))
        env.run(until=100)

    def source(self, env, num_people, mean_arrival, service_time):
        for i in range(num_people):
            yield env.timeout(rand.expovariate(1 / mean_arrival))
            env.process(self.person(env, i + 1, service_time))

    def person(self, env, id, service_time):
        arrival_time = env.now
        print(f"Person {id} arrives at {arrival_time}")

        # Elevator arrives after some time
        wait = rand.expovariate(1 / service_time)
        yield env.timeout(wait)

        print(f"Person {id} got elevator at {env.now}")


if __name__ == '__main__':
    sim = ElevatorEventSim()
    sim.runsim(10, 5, 3)