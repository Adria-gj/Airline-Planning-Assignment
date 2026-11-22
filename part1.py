# #=============================================================================================================================
# # AE4423 Airline Planning & Optimisation - Assignment Part 1, Question 1A
# # Network & Fleet Development (route-based hub-and-spoke model with triangular routes and electric aircraft)
# # Group 17 - Southern Star (hub: Madrid)
# # Names / student IDs: ...
# #=============================================================================================================================
from gurobipy import *
from openpyxl import *
from time import *
from math import *
from itertools import permutations


#=============================================================================================================================
#  Parameters (Appendices A, B, D)
#=============================================================================================================================
LF        = 0.8      # maximum average load factor
FUEL      = 1.42     # fuel price [EUR/gallon]
ENERGY    = 0.07     # energy price [EUR/kWh]
DAYS      = 7        # one week of operations

#              speed seats TAT charge range runway lease   CX    CT    CF   battery BT  electric
AC_DATA = {1: (550,  45,   25,  0,    1500, 1400,  15000,  300,  750, 1.0,     0,  12, False),
           2: (820,  70,   35,  0,    3300, 1600,  34000,  600,  775, 2.0,     0,  12, False),
           3: (850, 150,   45,  0,    6300, 1800,  80000, 1250, 1400, 3.75,    0,  10, False),
           4: (350,  20,   20, 20,     400,  750,  10000,   90,  750, 0.0,  2130,   9, True),
           5: (480,  48,   25, 45,    1000,  950,  16000,  120,  750, 0.0,  8216,   9, True)}


#=============================================================================================================================
#  Classes for airports, aircraft types and routes
#=============================================================================================================================
class Airport:
    def __init__(self, code, lat, lon, runway, slots):
        self.code   = code
        self.lat    = lat
        self.lon    = lon
        self.runway = runway
        self.slots  = slots          # weekly landing slots (None = unlimited, hub)

class Aircraft:
    def __init__(self, k, data):
        (self.speed, self.seats, self.TAT, self.charge, self.range, self.runway, self.lease,
         self.CX, self.CT, self.CF, self.battery, self.BT, self.electric) = data
        self.id = k

class Route:
    def __init__(self, nodes):
        self.nodes  = nodes                              # [H,i,H] or [H,i,j,H]
        self.legs   = list(zip(nodes[:-1], nodes[1:]))   # [(H,i),(i,H)] or [(H,i),(i,j),(j,H)]
        self.spokes = nodes[1:-1]                        # airports where the route lands (uses slots)
        self.cost   = {}                                 # cost per rotation for each aircraft type
        self.time   = {}                                 # time per rotation for each aircraft type
        self.types  = []                                 # aircraft types that can fly this route

    def name(self):
        return '-'.join(self.nodes)

    def legs_used(self, origin, destination):
        """Indices of the legs a passenger uses travelling origin -> destination on this route.
        The hub appears at the start AND the end: as an origin we use the start, as a destination the end."""
        a = 0 if origin == self.nodes[0] else self.nodes.index(origin)
        b = len(self.nodes) - 1 if destination == self.nodes[-1] else self.nodes.index(destination)
        return list(range(a, b)) if a < b else []


#=============================================================================================================================
#  Auxiliary functions
#=============================================================================================================================
def distance(A, B):
    """Great-circle distance in km (Appendix C)."""
    p1, p2, l1, l2 = radians(A.lat), radians(B.lat), radians(A.lon), radians(B.lon)
    return 6371 * 2 * asin(sqrt(sin((p1 - p2) / 2)**2 + cos(p1) * cos(p2) * sin((l1 - l2) / 2)**2))


def read_data(file):
    """Read hub, airports and weekly demand from the group Excel file."""
    wb   = load_workbook(file, data_only=True)
    rows = [r for r in wb.active.iter_rows(values_only=True) if any(c is not None for c in r)]

    cities, codes = rows[2][2:], rows[3][2:]
    hub = codes[cities.index(rows[1][2])]                # hub city given in row 2 ('Madrid')

    Airports = {}
    for n, c in enumerate(codes):
        slots = None if c == hub else int(rows[7][2 + n])
        Airports[c] = Airport(c, rows[4][2 + n], rows[5][2 + n], rows[6][2 + n], slots)

    q = {}
    for r in rows[9:9 + len(codes)]:
        for n, dest in enumerate(codes):
            q[r[1], dest] = r[2 + n]
    return hub, Airports, q


def leg_cost(ac, i, j, d, hub):
    """Operating cost of one flight leg (Appendix B). 30% discount on CX, CT, CF if the leg touches the hub."""
    cost   = ac.CX + ac.CT * d[i, j] / ac.speed + ac.CF * FUEL / 1.5 * d[i, j]
    energy = ENERGY * ac.battery * d[i, j] / ac.range if ac.electric else 0
    return (0.7 * cost if hub in (i, j) else cost) + energy


def leg_time(ac, i, j, d, hub):
    """Flight time + turnaround at arrival [h]. Kerosene: TAT x1.5 into the hub. Electric: + charging time at the hub."""
    if ac.electric:
        tat = ac.TAT + (ac.charge if j == hub else 0)
    else:
        tat = ac.TAT * (1.5 if j == hub else 1)
    return d[i, j] / ac.speed + tat / 60


def generate_routes(hub, Airports, Aircrafts, d):
    """Task 1: all H-i-H and H-i-j-H routes, and the aircraft types that can fly each of them."""
    spokes = [a for a in Airports if a != hub]
    all_routes = [Route([hub, i, hub]) for i in spokes] + \
                 [Route([hub, i, j, hub]) for i, j in permutations(spokes, 2)]

    Routes = []
    for r in all_routes:
        for k, ac in Aircrafts.items():
            if ac.electric:                                     # electric: whole rotation on one charge
                range_ok = sum(d[a, b] for a, b in r.legs) <= ac.range
            else:                                               # kerosene: every leg within range
                range_ok = max(d[a, b] for a, b in r.legs) <= ac.range
            runway_ok = all(Airports[s].runway >= ac.runway for s in r.spokes)
            time      = sum(leg_time(ac, a, b, d, hub) for a, b in r.legs)
            time_ok   = time <= ac.BT                            # one rotation fits in one day of block time
            if range_ok and runway_ok and time_ok:
                r.types.append(k)
                r.time[k] = time
                r.cost[k] = sum(leg_cost(ac, a, b, d, hub) for a, b in r.legs)
        if r.types:
            Routes.append(r)
    return Routes


def check_model_status(model):
    status = model.status
    if status != GRB.Status.OPTIMAL:
        if status == GRB.Status.INFEASIBLE:
            print('The model is infeasible; computing IIS')
            model.computeIIS()
            for c in model.getConstrs():
                if c.IISConstr:
                    print(c.constrName)
            exit(0)
        print('Optimization was stopped with status', status)


#=============================================================================================================================
#  Optimization of the model through Gurobi
#=============================================================================================================================
def Network_Fleet_Problem(hub, Airports, Aircrafts, Routes, q, d, Y):

    model  = Model("Network_Fleet")
    spokes = [a for a in Airports if a != hub]
    R      = range(len(Routes))

    # ---------------- Decision variables ----------------
    z  = {}   # z[r,k]      weekly frequency of route r with aircraft type k              (integer)
    AC = {}   # AC[k]       number of aircraft of type k leased                           (integer)
    x  = {}   # x[i,j,r]    passengers i->j on route r without connecting at the hub
    w  = {}   # w[i,j]      passengers i->j connecting at the hub (as in Lecture 3)
    wi = {}   # wi[i,j,r]   connecting passengers i->j flying their first flight (i->hub) on route r
    wo = {}   # wo[i,j,r]   connecting passengers i->j flying their second flight (hub->j) on route r

    for k in Aircrafts:
        AC[k] = model.addVar(vtype=GRB.INTEGER, name=f'AC({k})')
    for r in R:
        for k in Routes[r].types:
            z[r, k] = model.addVar(vtype=GRB.INTEGER, name=f'z({Routes[r].name()},{k})')

    for r in R:
        for i in Routes[r].nodes:
            for j in Routes[r].nodes:
                if i != j and q[i, j] > 0 and Routes[r].legs_used(i, j):
                    x[i, j, r] = model.addVar(name=f'x({i},{j},{r})')

    for i in spokes:
        for j in spokes:
            if i != j and q[i, j] > 0:
                w[i, j] = model.addVar(name=f'w({i},{j})')
                for r in R:
                    if i in Routes[r].spokes:
                        wi[i, j, r] = model.addVar(name=f'wi({i},{j},{r})')
                    if j in Routes[r].spokes:
                        wo[i, j, r] = model.addVar(name=f'wo({i},{j},{r})')
    model.update()

    # ---------------- Objective: maximise weekly profit ----------------
    revenue  = quicksum(Y[i, j] * d[i, j] * x[i, j, r] for (i, j, r) in x) \
             + quicksum(0.8 * Y[i, j] * d[i, j] * w[i, j] for (i, j) in w)        # transfer pax pay 20% less
    op_costs = quicksum(Routes[r].cost[k] * z[r, k] for (r, k) in z)
    leasing  = quicksum(Aircrafts[k].lease * AC[k] for k in Aircrafts)
    model.setObjective(revenue - op_costs - leasing, GRB.MAXIMIZE)

    # ---------------- C1: demand ----------------
    for (i, j) in q:
        if i != j and q[i, j] > 0:
            model.addConstr(quicksum(x[i, j, r] for r in R if (i, j, r) in x)
                            + (w[i, j] if (i, j) in w else 0) <= q[i, j], name=f'Demand({i},{j})')

    # ---------------- C1*: every connecting passenger flies one flight into and one flight out of the hub --------------
    for (i, j) in w:
        model.addConstr(quicksum(wi[i, j, r] for r in R if (i, j, r) in wi) == w[i, j], name=f'TransferIn({i},{j})')
        model.addConstr(quicksum(wo[i, j, r] for r in R if (i, j, r) in wo) == w[i, j], name=f'TransferOut({i},{j})')

    # ---------------- C2: capacity on every leg of every route ----------------
    for r in R:
        route = Routes[r]
        seats = quicksum(Aircrafts[k].seats * LF * z[r, k] for k in route.types)
        for l, (a, b) in enumerate(route.legs):
            pax  = quicksum(x[i, j, s] for (i, j, s) in x  if s == r and l in route.legs_used(i, j))
            pax += quicksum(wi[i, j, s] for (i, j, s) in wi if s == r and l in route.legs_used(i, hub))
            pax += quicksum(wo[i, j, s] for (i, j, s) in wo if s == r and l in route.legs_used(hub, j))
            model.addConstr(pax <= seats, name=f'Capacity({route.name()},{a}-{b})')

    # ---------------- C3: flow balance -> not needed, every rotation starts and ends at the hub ----------------

    # ---------------- C4: aircraft productivity (weekly block hours) ----------------
    for k, ac in Aircrafts.items():
        model.addConstr(quicksum(Routes[r].time[k] * z[r, k] for r in R if (r, k) in z)
                        <= ac.BT * DAYS * AC[k], name=f'BlockTime({k})')

    # ---------------- C5: range & runway -> handled in generate_routes (z only exists for feasible routes) --------------

    # ---------------- C6: weekly landing slots at the spokes ----------------
    for s in spokes:
        model.addConstr(quicksum(z[r, k] for (r, k) in z if s in Routes[r].spokes)
                        <= Airports[s].slots, name=f'Slots({s})')

    model.setParam("LogFile", 'log_file')
    model.setParam("TimeLimit", 900)
    model.update()
    model.write("NetworkFleet_Model.lp")
    model.optimize()

    check_model_status(model)
    print_model_result(model, hub, Aircrafts, Routes, z, AC, x, wi, wo, w, d)


#=============================================================================================================================
#  Print results (Task 5)
#=============================================================================================================================
def print_model_result(model, hub, Aircrafts, Routes, z, AC, x, wi, wo, w, d):
    print('\nWeekly profit = %.0f EUR' % model.ObjVal)
    print('Passengers: direct %.0f, transfer %.0f' % (sum(v.X for v in x.values()), sum(v.X for v in w.values())))

    print('\nFleet:')
    for k, ac in Aircrafts.items():
        n = round(AC[k].X)
        if n > 0:
            used = sum(Routes[r].time[k] * z[r, k].X for (r, kk) in z if kk == k)
            print(f'  AC{k}: {n} aircraft, utilisation {used / (ac.BT * DAYS * n) * 100:.1f}%')

    print('\nNetwork:')
    for (r, k), var in z.items():
        if var.X > 0.5:
            route = Routes[r]
            seats = sum(Aircrafts[kk].seats * z[r, kk].X for kk in route.types)
            lf = []
            for l, (a, b) in enumerate(route.legs):
                pax  = sum(x[i, j, s].X for (i, j, s) in x if s == r and l in route.legs_used(i, j))
                pax += sum(wi[i, j, s].X for (i, j, s) in wi if s == r and l in route.legs_used(i, hub))
                pax += sum(wo[i, j, s].X for (i, j, s) in wo if s == r and l in route.legs_used(hub, j))
                lf.append(f'{a}-{b}: {pax / seats * 100:.0f}%')
            print(f'  {route.name():<24} AC{k}  {round(var.X):2d}/week   LF {", ".join(lf)}')


#=============================================================================================================================
#  Run the program
#=============================================================================================================================
if __name__ == '__main__':
    hub, Airports, q = read_data("DemandGroup17.xlsx")
    Aircrafts = {k: Aircraft(k, AC_DATA[k]) for k in AC_DATA}

    d = {(i, j): distance(Airports[i], Airports[j]) for i in Airports for j in Airports if i != j}
    Y = {(i, j): 5.9 * d[i, j]**-0.76 + 0.043 for (i, j) in d}       # yield [EUR/RPK] (Appendix A)

    Routes = generate_routes(hub, Airports, Aircrafts, d)
    print(f'{len(Routes)} feasible routes generated')

    start_time = time()
    Network_Fleet_Problem(hub, Airports, Aircrafts, Routes, q, d, Y)
    print('Run Time =', time() - start_time)
