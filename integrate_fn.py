
print("Which state of water you needs, Saturated state or Non - Saturated state (Subcooled/Superheated)")
print("0 = Non - Saturated state (Subcooled/Superheated), 1 = Saturated state, 2 = Reverse Calculation")
try:
    sat = int(input('pls insert your case of state : = '))
    sat = str(sat)
    if sat == '1': 
        print('your scenario = Saturated')
        from sat_fn import sat_calc as fx
    elif sat == '0': 
        print('your scenario = Non - Saturated')
        from cleaner_interpolate_fn import non_sat_fn as fx
    elif sat == '2' :
        print('your scenario = Reverse Calculation')
        from rev_calc import rev_fn as fx
    fx()
except ValueError:
    print('Invalid value, please run the program again and insert the new one.')
except:
    print('Something else went wrong.')

