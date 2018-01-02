import xlsxwriter as xls

class ExcelWriter:
    def __init__(self, sim):
        self.sim = sim
        self.workbook = xls.Workbook(sim.output)
        self.abs = self.workbook.add_worksheet("Absolute")
        self.perc = self.workbook.add_worksheet("Percentage")
        self.surv = self.workbook.add_worksheet("Survival")
        self.charts = self.workbook.add_worksheet("Chart")

    def generate(self):
        self.surv.write(0,0,"Generation")
        self.surv.write(1,0,"Count")
        self.surv.write(2,0,"Remaining")
        self.surv.write(3,0,"Probability of survival")
        for i, (gen, count) in enumerate(zip(self.sim.generations, self.sim.count)):
            self.abs.write(0, i, i)
            self.perc.write(0, i, i)
            self.surv.write(0, i+1, i)
            self.surv.write(1, i+1, self.sim.count[i])
            self.surv.write(2, i+1, self.sim.surv[i])
            self.surv.write(3, i+1, self.sim.surv[i]/self.sim.n)
            for j, val in enumerate(gen):
                self.abs.write(j+1, i, val)
                try:
                    self.perc.write(j+1, i, val/count)
                except ZeroDivisionError:
                    self.perc.write(j+1, i, 0)

        stacked_chart = self.workbook.add_chart({'type': 'column', 'subtype': 'percent_stacked'})
        max_row = len(self.sim.generations)
        for col in range(self.sim.n):
            stacked_chart.add_series({
                #'name': ['Percentage', 0, col_num],
                #'categories' : ['Percentage', 1, 0, max_row, 0],
                'values': ['Percentage', 1, col, max_row, col],
                'gap': 0
            })

        stacked_chart.set_size({'width': 1920, 'height': 1080})
        stacked_chart.set_legend({'none': True})
        self.charts.insert_chart("A1", stacked_chart)
        stacked_chart.set_y_axis({'visible': False, 'major_gridlines': {'visible': False}})
        stacked_chart.set_x_axis({'visible': False, 'major_gridlines': {'visible': False}})



        self.workbook.close()

        print("Done")
