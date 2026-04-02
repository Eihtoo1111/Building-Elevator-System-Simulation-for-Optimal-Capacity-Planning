import random
import matplotlib.pyplot as plt

rand = random.Random()

class DiscreteTimeSim:

    def __init__(self):
        self.t = 0
        self.interval = 1
        self.line = 0
        self.timesteps = []
        self.linesizes = []

    def initialize(self, timeinterval):
        self.t = 0
        self.interval = timeinterval
        self.line = 0
        self.timesteps = [self.t]
        self.linesizes = [self.line]

    def observe(self):
        self.timesteps.append(self.t)
        self.linesizes.append(self.line)

    def update(self):
        # Random arrivals
        if rand.random() < 0.3:
            self.line += 1

        # Service one person
        if self.line > 0 and rand.random() < 0.5:
            self.line -= 1

        self.t = self.t + self.interval

    def runsim(self, timeinterval, endtime):
        self.initialize(timeinterval)

        while self.t < endtime:
            self.update()
            self.observe()

        plt.plot(self.timesteps, self.linesizes)
        plt.xlabel("Time")
        plt.ylabel("Line Size")
        plt.show()


if __name__ == '__main__':
    sim = DiscreteTimeSim()
    sim.runsim(1, 100)