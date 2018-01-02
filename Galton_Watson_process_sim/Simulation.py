import numpy
from .ExcelWriter import ExcelWriter


class Simulation():
    def __init__(self, initial, l, m, output, steps, graph_type):
        self.n = initial
        self.l = l
        self.m = m
        self.output = output
        self.steps = steps

        self.generations = list()
        self.count = [initial]
        self.surv = [initial]
        self.generations.append([1]*self.n)

        self.excelWriter = ExcelWriter(self, graph_type)

    def simulate(self):
        prev = self.n
        same = 0
        gen = 0
        while (self.steps > 0 and same < self.steps) or (self.steps == 0 and gen < self.m):
            self.generations.append([0]*self.n)
            count = 0
            surv = 0
            for i in range(self.n):
                next_gen = sum(numpy.random.poisson(self.l, self.generations[gen][i]))
                count += next_gen
                if next_gen > 0:
                    surv += 1
                self.generations[gen+1][i] = next_gen
            self.count.append(count)
            self.surv.append(surv)

            if self.steps > 0:
                if surv != prev:
                    same = 0
                else:
                    same += 1

                prev = surv
                if same >= self.steps:
                    break


            #print("gen:",gen,":",surv)
            #print(count)
            gen += 1
        print("Simulation completed\nSaving results")
        self.excelWriter.generate()

