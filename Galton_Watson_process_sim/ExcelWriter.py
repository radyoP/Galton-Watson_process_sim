import xlsxwriter as xls
import numpy as np
import matplotlib.pyplot as plt


class ExcelWriter:
    def __init__(self, sim, graph_type):
        self.sim = sim
        self.workbook = xls.Workbook(sim.output)
        self.abs = self.workbook.add_worksheet("Absolute")
        self.surv = self.workbook.add_worksheet("Survival")
        self.charts = self.workbook.add_worksheet("Chart")
        self.graph_type = graph_type

    def generate(self):

        print("Filling spreadsheets")
        self.fill_sheets()

        print("Creating small charts")
        self.create_small_charts()

        if self.sim.n < 256:
            print("Creating large chart")
            self.create_big_chart()
            self.workbook.close()

        else:
            print("Maximum number of entities Excel can handle was exceeded. Large chart won't be complete")
            self.create_big_chart()
            self.workbook.close()
            print("Attempting to create chart using matplotlip. This can take a while, but .xlsx file is finished")
            print("If you don't want to wait, you can close this program")
            self.create_big_plt()


        print("\nDone")


    def fill_sheets(self):
        """Fills all the sheets with data"""

        self.surv.write(0, 0, "Generation")
        self.surv.write(1, 0, "Count")
        self.surv.write(2, 0, "Remaining")
        self.surv.write(3, 0, "Probability of survival")

        for i in range(self.sim.n):
            self.abs.write(i+1, 0, i)

        for i, (gen, count) in enumerate(zip(self.sim.generations, self.sim.count)):
            self.abs.write(0, i+1, i)
            self.surv.write(0, i+1, i)
            self.surv.write(1, i+1, self.sim.count[i])
            self.surv.write(2, i+1, self.sim.surv[i])
            self.surv.write(3, i+1, self.sim.surv[i]/self.sim.n)
            for j, val in enumerate(gen):
                self.abs.write(j+1, i+1, val)

        per_format = self.workbook.add_format(({'num_format': '0%'}))
        self.surv.set_row(3, None, per_format)

    def create_small_charts(self):
        max_col = len(self.sim.generations)

        # Count Chart
        count_chart = self.workbook.add_chart({'type': 'scatter', 'subtype': 'straight'})
        count_chart.add_series({
            'name': ['Survival', 1, 0],
            'categories': ['Survival', 0, 1, 0, max_col],
            'values': ['Survival', 1, 1, 1, max_col]
        })
        count_chart.set_legend({'none': True})
        self.surv.insert_chart("A6", count_chart)

        # Remaining Chart
        rem_chart = self.workbook.add_chart({'type': 'scatter', 'subtype': 'straight'})
        rem_chart.add_series({
            'name': ['Survival', 1, 0],
            'categories': ['Survival', 0, 2, 0, max_col],
            'values': ['Survival', 2, 1, 2, max_col]
        })
        rem_chart.set_legend({'none': True})
        self.surv.insert_chart("I6", rem_chart)

        # Probability of survival chart
        surv_prob_chart = self.workbook.add_chart({'type': 'scatter', 'subtype': 'straight'})
        surv_prob_chart.add_series({
            'name': ['Survival', 3, 0],
            'categories': ['Survival', 0, 1, 0, max_col],
            'values': ['Survival', 3, 1, 3, max_col]
        })
        surv_prob_chart.set_legend({'none': True})
        self.surv.insert_chart("Q6", surv_prob_chart)

    def create_big_chart(self):
        max_col = len(self.sim.generations)

        stacked_chart = self.workbook.add_chart(self.graph_type)

        for row in range(1, self.sim.n):
            stacked_chart.add_series({
                'name': ['Absolute', row, 0],
                'values': ['Absolute', row, 1, row, max_col],
                'gap': 0
            })

        stacked_chart.set_size({'width': 1920, 'height': 1080})
        stacked_chart.set_legend({'none': True})
        self.charts.insert_chart("A1", stacked_chart)
        stacked_chart.set_y_axis({'visible': True, 'major_gridlines': {'visible': False}})
        stacked_chart.set_x_axis({'visible': True, 'major_gridlines': {'visible': False}})

    def create_big_plt(self):
        N = len(self.sim.generations)
        ind = np.arange(N)
        width = 1
        for gen, count in zip(self.sim.generations, self.sim.count):
            for i, value in enumerate(gen):
                gen[i] /= count
        data = np.array(self.sim.generations).T
        bottom_size = [0] * N

        plt.figure(figsize=(18.0, 12.0))

        print("Creating plot")
        for i in range(0, self.sim.n):

            plt.bar(ind, data[i], width, bottom=bottom_size)
            bottom_size = [bottom_size[j] + data[i][j] for j in range(N)]

        print("Plot finished, saving image")

        if(self.sim.steps == 0):
            plt.title("Initial: {}, lambda: {}, stopping criteria: Max steps {}".format(self.sim.n, self.sim.l, self.sim.m))
        else:
            plt.title(
                "Initial: {}, lambda: {}, stopping criteria: Same steps {}".format(self.sim.n, self.sim.l, self.sim.steps))


        plt.savefig(self.sim.output[:-5]+".png")
