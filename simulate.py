import argparse
from Simulation import Simulation

parser = argparse.ArgumentParser(description="Starts the simulation of Galton-Watson process")

parser.add_argument("-n","--initial", type=int, default="255", help="Number of initial entities")
parser.add_argument("-l", "--LAMBDA", type=float, default="1.1", help="Lambda of Poisson distribution .")
parser.add_argument("-o", "--output", type=str, default="def", help="Output file")
group = parser.add_mutually_exclusive_group()
group.add_argument("-s", "--same_steps", type=int, default=0, help="Simulation ends, if number of remaining entities "
                                                                   "have not decreased in given number of steps")
group.add_argument("-m", "--max_steps", type=int, default="100", help="Simulation ends, in given number of steps")

chart_type = parser.add_mutually_exclusive_group()
chart_type.add_argument("-c", "--column", dest="type", action="store_const", const={'type': 'column', 'subtype': 'stacked'}, help="Generates column stacked chart")
chart_type.add_argument("-cp", "--column_percentage", dest="type", action="store_const", const={'type': 'column', 'subtype': 'percent_stacked'}, help="Generates column percent stacked chart")
chart_type.add_argument("-a", "--area", dest="type", action="store_const", const={'type': 'area', 'subtype': 'stacked'}, help="Generates area stacked chart")
chart_type.add_argument("-ap", "--area_percentage", dest="type", action="store_const", const={'type': 'area', 'subtype': 'percent_stacked'}, help="Generates area percent stacked chart")


args = parser.parse_args()

if args.type == None:
    args.type = {'type': 'column', 'subtype': 'percent_stacked'}
if args.initial < 1:
    print("Number of initial surnames must be greater than 0")
    exit(1)
if args.max_steps < 1:
    print("Max steps must be greater than 0")
    exit(1)
if args.LAMBDA < 0:
    print("Lambda must be greater than 0")
    exit(1)
if args.same_steps < 0:
    print("Same steps must be greater than 0")

if args.same_steps > 0:
    if args.output == "def":
        args.output = "sim_n-{}_l-{}_same_steps-{}.xlsx".format(args.initial, args.LAMBDA, args.same_steps)
    print("Starting simulation\n\nInitial number: {}\nStopping Criteria - Same steps: {}\nLambda: {}\nOutput file: {}\nChart type: {} {}".format(
        args.initial, args.same_steps, args.LAMBDA, args.output, args.type['type'], args.type['subtype']))
else:
    if args.output == "def":
        args.output = "sim_n-{}_l-{}_max_steps-{}.xlsx".format(args.initial, args.LAMBDA, args.max_steps)
    print("Starting simulation\n\nInitial number: {}\nStopping Criteria - Maximum criteria: {}\nLambda: {}\nOutput file: {}\nChart type: {} {}".format(args.initial,args.max_steps, args.LAMBDA, args.output, args.type['type'], args.type['subtype']))


sim = Simulation(args.initial, args.LAMBDA, args.max_steps, args.output, args.same_steps, args.type)
sim.simulate()