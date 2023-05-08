
# =============================================================================#

def non_sat_fn():
    import numpy as np
    import math
    from openpyxl import load_workbook

    def Valid_TandP(t, p):
        inRange_T = 10 <= t <= 1000
        inRange_P = 0.10 <= p <= 150
        if not inRange_P and not inRange_T: return False, 'both of your inputs are not in condition range.'
        elif not inRange_T: return False, 'your temperature is not in condition range.'
        elif not inRange_P: return False, 'your pressure is not in condition range.'
        return True, 'You are all set !!'

    def nearestPressure(ref_P, P):
        # if some of ref_P have the same diff. from P, argmin() always return the less one.
        indexOfMinP = np.abs(ref_P - P).argmin()
        return ref_P[indexOfMinP], indexOfMinP

    def nearestTemperature(ref_T, T):
        indexOfMinP = nearestPressure(ref_P, P)[1]
        indexOfMinT = np.abs(ref_T[indexOfMinP] - T).argmin()
        return ref_T[indexOfMinP, indexOfMinT], [indexOfMinP, indexOfMinT]

    def cellChecker(fn):
        # row = fn[0] # 0 - 28
        # col = fn[1] # 0 - 14
        row, col = fn
        if row in [0, 29] and col in [0, 14]: return 'Corner'
        elif row in [0, 29] or col in [0, 14]: return 'Neither'
        else: return 'Regular'

    def Nz(First, Second):
        return abs((Second/First)-1)

    def calcTriangle(rows, cols):
        result = [[Nz(T, ref_T[i, j]), Nz(P, ref_P[i]), (i, j)] for i in rows for j in cols]
        return [(ref_P[res[2][0]], ref_T[res[2]], res[2]) for res in sorted(result)[:3]]

    def secondNearestT(fn, pattern):
        if pattern == 'Regular':

            x, y = [e-1 for e in fn] 
            if T < value1st_T: cols = np.arange(y, y+2)
            elif T > value1st_T: cols = np.arange(y+1, y+3)
            elif T == value1st_T and P < value1st_P: cols = np.arange(y, y+2)
            elif T == value1st_T and P > value1st_P: cols = np.arange(y+1, y+3)
            if P < value1st_P: rows = np.arange(x, x+2)
            elif P > value1st_P: rows = np.arange(x+1, x+3)
            elif P == value1st_P and T < value1st_T: rows = np.arange(x, x+2)
            elif P == value1st_P and T > value1st_T: rows = np.arange(x+1, x+3)

        elif pattern == 'Neither':

            x, y = fn

            if x == 0: rows = np.arange(2)
            elif x == 29: rows = np.arange(-2, 0)
            elif P > value1st_P: rows = np.arange(x, x+2)
            elif P < value1st_P: rows = np.arange(x-1, x+1)
            elif P == value1st_P and T < value1st_T: rows = np.arange(x-1, x+1)
            elif P == value1st_P and T > value1st_T: rows = np.arange(x, x+2)

            if y == 0: cols = np.arange(2)
            elif y == 14: cols = np.arange(-2, 0)
            elif T > value1st_T: cols = np.arange(y, y+2)
            elif T < value1st_T: cols = np.arange(y-1, y+1)
            elif T == value1st_T and P < value1st_P: cols = np.arange(y-1, y+1)
            elif T == value1st_T and P > value1st_P: cols = np.arange(y, y+2)

        elif pattern == 'Corner':

            x, y = fn
            if x == 0: rows = np.arange(2)
            elif x == 29: rows = np.arange(-2,0)

            if y == 0: cols = np.arange(2)
            elif y == 14: cols = np.arange(-2, 0)

        return calcTriangle(rows, cols)
     
    def solVar(ma, mb):
        a = np.array(ma)
        b = np.array(mb)
        x = np.linalg.solve(a, b)
        return x

#=================================================================================#

    book = load_workbook(filename="C:\\Users\\spypl\\OneDrive\\เดสก์ท็อป\\work\\p and t (version 1).xlsx")
    sheet = book["y"]
    print("please insert your data\ntemperature range = 10 - 1000 °C\npressure range = 0.10 - 150.00 Bar")
    T = float(input("Temperature (°C) = "))
    P = float(input("Pressure (Bar) = "))
    # when a function return more than one value, you can collect it by declare variables w/
    # !! Elements wise !! (as down below)
    valid, err = Valid_TandP(T, P)
    if not valid: 
        print(err)
        return # return for stopping the non_sat_fn() function when T and/or P not valid
    ref_P = np.array([cell.value for col in sheet["A2":"A31"] for cell in col])
    ref_T = np.array([cell.value for row in sheet["B2":"P31"] for cell in row]).reshape((30,15))
    r = np.array([cell.value for row in sheet["B34":"Q63"] for cell in row]).reshape((30,16))
    h = np.array([cell.value for row in sheet["B66":"Q95"] for cell in row]).reshape((30,16))
    s = np.array([cell.value for row in sheet["B98":"Q127"] for cell in row]).reshape((30,16))
    if T in ref_T[:, 7]:
        print('T state = Saturated, please use another condition.')
        return 
    if T in ref_T and P in ref_P:
        print('your P and T are already in the table')
        pi = np.where(ref_P == P)[0][0]
        ti = np.where(ref_T[pi] == T)[0][0]
        if ti != 0: ti += 1
        x = (pi, ti)
        print(f"r = {r[x]} kg/m^3,\nh = {h[x]} kJ/kg,\ns = {s[x]} kJ/(kg K)")
        return    
    value1st_T, index1st_T = nearestTemperature(ref_T, T) # index1st_T is abbreviated as fn
    value1st_P, index1st_P = nearestPressure(ref_P, P)
    chosen = secondNearestT(index1st_T, cellChecker(index1st_T)) # [(P, T, id),..]
    pp = list(set([e[2][0] for e in chosen]))
    T_sat_max = max([ref_T[i, 7] for i in pp])
    ma = [[math.log(e[1]), math.log(e[0]), 1] for e in chosen]
    mr = []
    mh = []
    ms = []
    for e in chosen:
        row, col = e[2]
        if T > T_sat_max: col += 1
        mr.append(float(math.log(r[row, col])))
        mh.append(float(math.log(h[row, col])))
        ms.append(float(math.log(s[row, col])))
    ar, br, cr = solVar(ma, mr)
    ah, bh, ch = solVar(ma, mh)
    az, bz, cz = solVar(ma, ms)
    cal_r = round(math.e**(ar*math.log(T)+br*math.log(P)+cr), 3)
    cal_h = round(math.e**(ah*math.log(T)+bh*math.log(P)+ch), 3)
    cal_s = round(math.e**(az*math.log(T)+bz*math.log(P)+cz), 3)
    state = 'Superheated' if T > ref_T[index1st_P, 7] else 'Subcooled'
    print('=== your calculated properties are right here !! ===')
    print(f"state = {state}.\nr = {cal_r} kg/m^3\nh = {cal_h} kJ/kg\ns = {cal_s} kJ/(kg K)")
    print('====================================================')
    return

#=================================================================================#

non_sat_fn()