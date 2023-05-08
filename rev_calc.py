def rev_fn():
    import numpy as np
    import math
    from openpyxl import load_workbook

    book = load_workbook(filename="C:\\Users\\spypl\\OneDrive\\เดสก์ท็อป\\work\\p and t (version 1).xlsx")
    sheet = book["y"]

    # Input & Validations
    print('this is rev_calc')
    print('pressure range = 0.1 - 150 Bar')
    p = float(input('pressure (bar) = '))
    if not (0.1 <= p <= 150): 
        print('pressure is invalid.')
        return
    print('r = density, h = enthalpy, s = entropy')
    # prop_name = input('property name = ')
    prop_name = 'r'
    if prop_name not in ['r', 'h', 's']:
        print('property name is invalid.')
        return
    if prop_name == 'r': 
        begin, stop = "B34", "Q63"
        
    elif prop_name == 'h': 
        begin, stop = "B66", "Q95"
        
    elif prop_name == 's': 
        begin, stop = "B98", "Q127"
        
    prop_val = float(input('property value = '))
    pv = np.array([cell.value for row in sheet[begin:stop] for cell in row]).reshape((30,16))
    mn = np.min(pv)
    mx = np.max(pv)
    if not (mn <= prop_val <= mx):
        print(f'property value is invalid. # not in range : {prop_val}, {prop_name}, p = {p} Bar')
        return
    
    # Calculations
    ref_P = np.array([cell.value for col in sheet["A2":"A31"] for cell in col])
    ref_T = np.array([cell.value for row in sheet["B2":"P31"] for cell in row]).reshape((30,15))
    pi = np.abs(ref_P - p).argmin() # if some of ref_P have the same diff. from P, argmin() always return the less one.
    floor = min(pv[pi, 7], pv[pi, 8])
    ceil = max(pv[pi, 7], pv[pi, 8])
    # Validation ( Special Case )
    if floor < prop_val < ceil :
        print('====================================================')
        print('State = Saturated')
        if p in ref_P: 
            print(f'Temperature = {ref_T[pi, 7]} °C')
            if prop_name == 'r':
                v = 1/prop_val
                vf = 1/pv[pi, 7]
                vg = 1/pv[pi, 8]
                x = (v - vf) / (vg - vf)
                xv = 1 / (x * vg + (1 -x) * vf)
            else: 
                sub_val = pv[pi, 7]
                sup_val = pv[pi, 8]
                x = (prop_val - pv[pi, 7]) / (pv[pi, 8] - pv[pi, 7])
                xv = x * sup_val + ( 1 - x ) * sub_val
        else: # if p do not already in the table > we have to interpolate
            coef = np.array([[4.60217915517292, 0.280826847523068, -0.0164154147623838, 0.00294857075682648, -0.000211754897215088],
                [6.86269032279953, -0.0257187997944348, -0.000630292934145276, 0.000813151631450814, -0.000635296986570567],
                [-0.522544165589532, 0.9434705944014, -0.0056138718033565, -0.00160797602409441, 0.00108138856513445],
                [6.03556649339902, 0.284662285786636, -0.0165034824738065, 0.00260465683673679, 4.04315225862806e-06],
                [7.89032264832234, 0.0142260966659471, 0.00212823596076632, 0.000502297527654881, -0.000321852266679515],
                [0.266334812987694, 0.244345324784465, -0.0187135540339003, 0.00254611173123604, -6.55800934700745e-05],
                [1.9948917567021, -0.0472976504911121, 0.000395514655242254, 0.000347131610971636, -0.000216532138758189]])
            T = round(math.e**(sum([coef[0, i] * math.log(p) ** i for i in range(5)])), 3)
            if prop_name == 'r': sub_id, sup_id = 1, 2
            elif prop_name == 'h': sub_id, sup_id = 3, 4
            elif prop_name == 's': sub_id, sup_id = 5, 6
            sub_val = round(math.e**(sum([coef[sub_id, i] * math.log(p) ** i for i in range(5)])), 3)
            sup_val = round(math.e**(sum([coef[sup_id, i] * math.log(p) ** i for i in range(5)])), 3)
            if prop_name == 'r':
                v = 1/prop_val
                vf = 1/sub_val
                vg = 1/sup_val
                x = (v - vf) / (vg - vf)
                xv = 1 / (x * vg + (1 -x) * vf)
            else: 
                x = (prop_val - sub_val) / (sup_val - sub_val)
                xv = x * sup_val + ( 1 - x ) * sub_val
            print(f'Temperature = {T} °C')
        print(f'x = {x}')
        # print(f'actual value = {prop_val}')
        # print(f'expected value = {xv}')
        print('====================================================')
        return
    vi = np.abs(pv[pi] - prop_val).argmin()
    if prop_val == pv[pi, vi] :
        if vi >= 8 : vi -= 1
        print('====================================================')
        print('This value is already in the database.')
        print('So, no calculations are needed, just retrieve the value.')
        print(f"Temperature = {ref_T[pi, vi]} °C")
        print('====================================================')
        return 
    fn = [pi, vi]
    def cellChecker(fn):
        # row = fn[0] # 0 - 29
        # col = fn[1] # 0 - 15
        row, col = fn
        if row in [0, 29] and col in [0, 15]: return 'Corner'
        elif row in [0, 29] or col in [0, 15]: return 'Neither'
        else: return 'Regular'
    
    pattern = cellChecker(fn)

    def Nz(First, Second):
        return abs((Second/First)-1)

    def calc(fn, pattern):

        value1st_P = ref_P[pi]
        prop_val_1st = pv[pi, vi]

        if pattern == 'Regular':

            x, y = [e-1 for e in fn] 
            if prop_name != 'r':
                if prop_val < prop_val_1st: cols = np.arange(y, y+2)
                elif prop_val > prop_val_1st: cols = np.arange(y+1, y+3)

            else:
                if prop_val < prop_val_1st: cols = np.arange(y+1, y+3)
                elif prop_val > prop_val_1st: cols = np.arange(y, y+2)
            if prop_val == prop_val_1st and p < value1st_P: cols = np.arange(y, y+2)
            elif prop_val == prop_val_1st and p > value1st_P: cols = np.arange(y+1, y+3)
            if p < value1st_P: rows = np.arange(x, x+2)
            elif p > value1st_P: rows = np.arange(x+1, x+3)
            elif p == value1st_P and prop_val < prop_val_1st: rows = np.arange(x, x+2)
            elif p == value1st_P and prop_val > prop_val_1st: rows = np.arange(x+1, x+3)

        elif pattern == 'Neither':

            x, y = fn

            if x == 0: rows = np.arange(2)
            elif x == 29: rows = np.arange(-2, 0)
            elif p > value1st_P: rows = np.arange(x, x+2)
            elif p < value1st_P: rows = np.arange(x-1, x+1)
            elif p == value1st_P and prop_val < prop_val_1st: rows = np.arange(x-1, x+1)
            elif p == value1st_P and prop_val > prop_val_1st: rows = np.arange(x, x+2)

            if prop_name == 'r':
                if prop_val < prop_val_1st: cols = np.arange(y, y+2)
                elif prop_val > prop_val_1st: cols = np.arange(y-1, y+1)
            else:
                if prop_val < prop_val_1st: cols = np.arange(y-1, y+1)
                elif prop_val > prop_val_1st: cols = np.arange(y, y+2)
            if y == 0: cols = np.arange(2)
            elif y == 15: cols = np.arange(-2, 0)

        elif pattern == 'Corner':

            x, y = fn
            if x == 0: rows = np.arange(2)
            elif x == 29: rows = np.arange(-2, 0)

            if y == 0: cols = np.arange(2)
            elif y == 15: cols = np.arange(-2, 0)

        result = [[Nz(prop_val, pv[i, j]), Nz(p, ref_P[i]), (i, j)] for i in rows for j in cols]
        # check index of T when state = Superheated 
        final = []
        # print([e[2] for e in sorted(result)[:3]])
        for res in sorted(result)[:3]:
            dup = []
            dup.append(ref_P[res[2][0]])
            if vi >= 8:
                m, n = res[2]
                dup.append(ref_T[m, n-1])
            else: dup.append(ref_T[res[2]])
            dup.append(res[2])
            final.append(dup)
        return final    
    
    chosen = calc(fn, pattern)
    ma = [[chosen[i][1], chosen[i][0], 1] for i in range(3)]
    mb = [float(pv[chosen[i][2]]) for i in range(3)]

    def solVar(ma, mb):
        a = np.array(ma)
        b = np.array(mb)
        x = np.linalg.solve(a, b)
        return x
    
    a, b, c = solVar(ma, mb)
    cal_T = round((prop_val - c - (b*p))/a, 3)
    print('=== your calculated temperature is right here !! ===')
    print(f"Temperature = {cal_T} °C")
    print('====================================================')
    return

rev_fn()