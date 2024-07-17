### QUESTION-2 ###

from datafile import *

### FORMULATION ###

# START TIME
start = time.time()
now = datetime.now()
print("\nStarted at -",now.strftime("%H:%M:%S"))

model = grb.Model()

### DECISION VARIABLES ###

QFinal = {}
DUnfulfilled = {}
DiscSD = {}
DiscQD = {}
DiscountSD = {}
DiscountQD = {}
Discount = {}

for i in range(num_items):
    DUnfulfilled[i] = model.addVar(lb=0, vtype=grb.GRB.CONTINUOUS, name=f'DUnfulfilled_{i}')

for v in range(num_vendors):
    DiscountSD[v] = model.addVar(lb=0, ub=1, vtype=grb.GRB.CONTINUOUS, name=f'DiscountSD_{v}')
    DiscountQD[v] = model.addVar(lb=0, ub=1, vtype=grb.GRB.CONTINUOUS, name=f'DiscountQD_{v}')
    Discount[v] = model.addVar(lb=0, ub=1, vtype=grb.GRB.CONTINUOUS, name=f'Discount_{v}')
    for i in range(num_items):
        QFinal[v, i] = model.addVar(lb=0, vtype=grb.GRB.CONTINUOUS, name=f'QFinal_{v}_{i}')
    for qd in range(num_QD):
        DiscQD[qd, v] = model.addVar(vtype=grb.GRB.BINARY, name=f'DiscQD_{qd}_{v}')
    for sd in range(num_SD):
        DiscSD[sd, v] = model.addVar(vtype=grb.GRB.BINARY, name=f'DiscSD_{sd}_{v}')

### CONSTRAINTS ###

# 1 - Quantity of each item taken from vendor can’t exceed quantity of that item with that vendor
for i in range(num_items):
    for v in range(num_vendors):
        model.addConstr(QFinal[v, i] <= Quantity[v][i])

# 2 - Demand balance for each item
for i in range(num_items):
    model.addConstr(grb.quicksum(QFinal[v, i] for v in range(num_vendors)) + DUnfulfilled[i] == Demand[i])

for v in range(num_vendors):
    
    # 5 - Total price of each vendor falls in any one slab of the spend discount slabs
    model.addConstr(grb.quicksum(DiscSD[sd, v] for sd in range(num_SD)) == 1)

    # 6 - Total quantity from each vendor falls in any one slab of the shipment discount slabs
    model.addConstr(grb.quicksum(DiscQD[qd, v] for qd in range(num_QD)) == 1)
    
    # 7 - Check in which spend discount slabs total price of each vendor lies
    # 7.1 - Spend discount slab 0 (<S1)
    model.addConstr(grb.quicksum(Price[v][i] * QFinal[v, i] for i in range(num_items)) - (S1 - E) <= M * (1 - DiscSD[0, v]))
    
    # 7.2 - Spend discount slab 1 (>=S1, <S2)
    model.addConstr(S1 - grb.quicksum(Price[v][i] * QFinal[v, i] for i in range(num_items)) <= M * (1 - DiscSD[1, v]))
    model.addConstr(grb.quicksum(Price[v][i] * QFinal[v, i] for i in range(num_items)) - (S2 - E) <= M * (1 - DiscSD[1, v]))

    # 7.3 - Spend discount slab 2 (>=S2)
    model.addConstr(S2 - grb.quicksum(Price[v][i] * QFinal[v, i] for i in range(num_items)) <= M * (1 - DiscSD[2, v]))

    # 8 - Check in which shipment discount slabs total quantity of each vendor lies
    # 8.1 - Shipment discount slab 0 (<Q1)
    model.addConstr(grb.quicksum(QFinal[v, i] for i in range(num_items)) - (Q1 - E) <= M * (1 - DiscQD[0, v]))

    # 8.2 - Shipment discount slab 1 (>=Q1, <Q2)
    model.addConstr(Q1 - grb.quicksum(QFinal[v, i] for i in range(num_items)) <= M * (1 - DiscQD[1, v]))
    model.addConstr(grb.quicksum(QFinal[v, i] for i in range(num_items)) - (Q2 - E) <= M * (1 - DiscQD[1, v]))

    # 8.3 - Shipment discount slab 2 (>=Q2)
    model.addConstr(Q2 - grb.quicksum(QFinal[v, i] for i in range(num_items)) <= M * (1 - DiscQD[2, v]))
    
    # 9 - Discount from spend discount criteria
    model.addConstr(DiscountSD[v] == DiscSD[0, v] * 0 + DiscSD[1, v] * x1 + DiscSD[2, v] * x2)

    # 10 - Discount from shipment discount criteria
    model.addConstr(DiscountQD[v] == DiscQD[0, v] * 0 + DiscQD[1, v] * y1 + DiscQD[2, v] * y2)

    # 11 - Total Discount
    model.addConstr(Discount[v] == (1 - DiscountSD[v]) * (1 - DiscountQD[v]))
    
### OBJECTIVE ###

cost = grb.quicksum(Discount[v] * Price[v][i] * QFinal[v, i] for v in range(num_vendors) for i in range(num_items)) + grb.quicksum(PenaltyCost[i] * DUnfulfilled[i] for i in range(num_items))

### SOLVER ###

model.setObjective(cost, sense = grb.GRB.MINIMIZE)
model.Params.NonConvex = 2 # Set the NonConvex parameter to 2
model.optimize()