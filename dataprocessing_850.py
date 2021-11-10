from openpyxl import load_workbook
from openpyxl.styles import Font
import math


def composition_profile(sheet):
    num_col = sheet["C"]
    repeat_col = sheet["D"]
    thick_col = sheet["F"]
    A_col = sheet["I"]
    B_col = sheet["J"]
    C_col = sheet["K"]
    notes_Col = sheet["Q"]

    num_list = []
    repeat_list = []
    thick_list = []
    A_list = []
    B_list = []
    C_list = []
    notes_list = []

    new_list_n = []
    new_list_A = []
    new_list_B = []
    new_list_C = []
    new_list = []

    for i in range(3,len(thick_col)):
        thick_list.append(thick_col[i].value)
    for j in range(3,len(repeat_col)):
        repeat_list.append(repeat_col[j].value)
    for k in range(3,len(num_col)):
        num_list.append(num_col[k].value)
    for m in range(3,len(A_col)):
        A_list.append(A_col[m].value)
    for o in range(3,len(B_col)):
        B_list.append(B_col[o].value)
    for p in range(3,len(C_col)):
        C_list.append(C_col[p].value)
    for n in range(3,len(notes_Col)):
        notes_list.append(notes_Col[n].value)
    
    

    index = 0
    while index < len(num_list):
        if num_list[index] == None:
            new_list.append(thick_list[index])
            new_list_A.append(A_list[index])
            new_list_B.append(B_list[index])
            new_list_C.append(C_list[index])
            new_list_n.append(notes_list[index])
            index += 1
        elif num_list[index] != None:
            repeat = repeat_list[index]

            get_repeat(num_list[index], repeat, new_list, thick_list, index)
            get_repeat(num_list[index], repeat, new_list_A, A_list, index)
            get_repeat(num_list[index], repeat, new_list_C, C_list, index)
            get_repeat(num_list[index], repeat, new_list_B, B_list, index)
            get_repeat(num_list[index], repeat, new_list_n, notes_list, index)

            index += num_list[index]
           
    return new_list, new_list_A, new_list_B, new_list_C, new_list_n

def print_A_profile(sheet,x_list,a_list):
    
    for m in range(0,len(x_list)):
        cell = sheet["J%d"%(m+3)]
        cell.value = x_list[m]
    
    for n in range(0,len(a_list)):
        cell = sheet["K%d"%(n+3)]
        cell.value = a_list[n]

def print_B_profile(sheet,x_list,a_list):
    
    for m in range(0,len(x_list)):
        cell = sheet["M%d"%(m+3)]
        cell.value = x_list[m]
    
    for n in range(0,len(a_list)):
        cell = sheet["N%d"%(n+3)]
        cell.value = a_list[n]

def print_C_profile(sheet,A_list,x_list,a_list):
    
    r = x_list[0]
    rows = A_list.index(r)
    # print(x_list[0])
    # print(rows)
    for m in range(0,len(x_list)):
        cell = sheet["P%d"%(m+3+rows)]
        cell.value = x_list[m]
    
    for n in range(0,len(a_list)):
        cell = sheet["Q%d"%(n+3+rows)]
        cell.value = a_list[n]

def A_Profile(sheet):
    x_col = sheet["C"]
    x_list = []
    new_x_list = []
    new_A_list = []
    A_col = sheet["D"]
    A_list = []
    notes_col = sheet["G"]
    notes_list = []

   
    for j in range(2,len(A_col)):
        if A_col[j].value != None:
            A_list.append(A_col[j].value)
            x_list.append(x_col[j].value)
            notes_list.append(notes_col[j].value)
   
    

    for i in range(len(A_list)-1):
        x1 = x_list[i]
        x2 = x_list[i+1]
        A = A_list[i+1]

        x = x1
        while x1 <= x < x2:
            if "CR" not in notes_list[i+1]:
                new_x_list.append(x)
                new_A_list.append(A)
                x += 0.1
            else:
                new_x_list.append(x)
                a = A_list[i] +(x-x1)*(A_list[i+1]-A_list[i])/(x2-x1)
                new_A_list.append(a)
                x += 0.1

    return new_x_list, new_A_list

def B_Profile(sheet):
    x_col = sheet["C"]
    x_list = []
    new_x_list = []
    new_B_list = []
    B_col = sheet["E"]
    B_list = []
    notes_col = sheet["G"]
    notes_list = []

   
    for j in range(2,len(B_col)):
        if B_col[j].value != None:
            B_list.append(B_col[j].value)
            x_list.append(x_col[j].value)
            notes_list.append(notes_col[j].value)
    
    for i in range(len(B_list)-1):
        x1 = x_list[i]
        x2 = x_list[i+1]
        B = B_list[i+1]

        x = x1
        while x1 <= x < x2:
            if "CR" not in notes_list[i+1]:
                new_x_list.append(x)
                new_B_list.append(B)
                x += 0.1
            else:
                new_x_list.append(x)
                a = B_list[i] +(x-x1)*(B_list[i+1]-B_list[i])/(x2-x1)
                new_B_list.append(a)
                x += 0.1

    return new_x_list, new_B_list

def C_Profile(sheet):
    x_col = sheet["C"]
    x_list = []
    new_x_list = []
    new_C_list = []
    C_col = sheet["F"]
    C_list = []
    notes_col = sheet["G"]
    notes_list = []

   
    for j in range(2,len(C_col)):
        if C_col[j].value != None:
            C_list.append(C_col[j].value)
            x_list.append(x_col[j].value)
            notes_list.append(notes_col[j].value)

    for i in range(len(C_list)-1):
        x1 = x_list[i]
        x2 = x_list[i+1]
        C = C_list[i+1]

        x = x1
        while x1 <= x < x2:
            if "CR" not in notes_list[i+1]:
                new_x_list.append(x)
                new_C_list.append(C)
                x += 0.1
            else:
                new_x_list.append(x)
                a = C_list[i] +(x-x1)*(C_list[i+1]-C_list[i])/(x2-x1)
                new_C_list.append(a)
                x += 0.1

    return new_x_list, new_C_list

def get_repeat(num, repeat, new_list, thick_list, index):
    for i in range (0, repeat):
        x = index
        for j in range (0, num):
            new_list.append(thick_list[x])
            x += 1

def print_profile_thick(sheet, cur_list,fname):
    
    for i in range(0,len(cur_list)):
        cell = sheet["B%d"%(i+4)]
        cell.value = cur_list[i]
    
    workbook.save(filename = fname)

def print_profile_A(sheet, cur_list,fname):
    
    for i in range(0,len(cur_list)):
        cell = sheet["D%d"%(i+4)]
        cell.value = cur_list[i]
    
    workbook.save(filename = fname)

def print_profile_B(sheet, cur_list, fname):
    
    for i in range(0,len(cur_list)):
        cell = sheet["E%d"%(i+4)]
        cell.value = cur_list[i]
    
    workbook.save(filename = fname)

def print_profile_C(sheet, cur_list,fname):
    
    for i in range(0,len(cur_list)):
        cell = sheet["F%d"%(i+4)]
        cell.value = cur_list[i]
    
    workbook.save(filename = fname)

def print_profile_notes(sheet, cur_list,fname):
    
    for i in range(0,len(cur_list)):
        cell = sheet["G%d"%(i+4)]
        cell.value = cur_list[i]
    
    workbook.save(filename = fname)

def print_profile_X(sheet, cur_list, fname):
    
    for i in range(0,len(cur_list)):
        cell = sheet["C%d"%(i+4)]
        cell.value = cur_list[i]
    
    workbook.save(filename = fname)

def print_E_profile(sheet, dist_sheet):

    E1_col = sheet["AE"]
    E2_col = sheet["AF"]
    E1_list = []
    E2_list = []
    for i in range (2, len(E1_col)):
        E1_list.append(E1_col[i].value)

    for j in range (2, len(E2_col)):
        E2_list.append(E2_col[j].value)   

    for m in range(0,len(E1_list)):
        cell = dist_sheet["S%d"%(m+3)]
        cell.value = E1_list[m]
    
    for n in range(0,len(E2_list)):
        cell = dist_sheet["T%d"%(n+3)]
        cell.value = E2_list[n]

def get_X(thick_list):
    X_list = []
    X_list.append(thick_list[0])
    for i in range(1, len(thick_list)):
        X_list.append(X_list[i-1]+thick_list[i])
    return X_list

def get_F_Dip(sheet):
    list_refl = []
    list_wave = [] 
    col_wave = sheet["A"]
    col_refl = sheet["B"]
    for i in range (4, len(col_refl)):
        if col_refl[i].value  != None:
            list_refl.append(col_refl[i].value)
        else:
            break
    for j in range (4, len(col_wave)):
        if col_wave[j].value  != None:
            list_wave.append(col_wave[j].value)
        else:
            break
    index1 = list_wave.index(820)
    index2 = list_wave.index(880)
    # print(index1,index2)
    # print(list_refl[index1])
    # print(list_refl[index2])
    
    list_temp = []
    for z in range(index1, index2+1):
        list_temp.append(list_refl[z])

    dip_refl = min(list_temp)
    dip_index = list_refl.index(dip_refl) #index
    dip_wave = list_wave[dip_index]
    
    return dip_wave

def get_B_Max(sheet,dip):
    list_refl = []
    list_wave = [] 
    list_trans = []
    col_wave = sheet["F"]
    col_refl = sheet["G"]
    col_trans = sheet["H"]
    for i in range (4, len(col_refl)):
        if col_refl[i].value  != None:
            list_refl.append(col_refl[i].value)
        else:
            break
    for j in range (4, len(col_wave)):
        if col_wave[j].value  != None:
            list_wave.append(col_wave[j].value)
        else:
            break
    for k in range (4, len(col_trans)):
        if col_trans[k].value  != None:
            list_trans.append(col_trans[k].value)
        else:
            break

    max_refl = max(list_refl)
    
    index_list = []
    for z in range(len(list_refl)):
        if list_refl[z] == max_refl:
            index_list.append(z)
    max_index = index_list[0]
    max_wave = list_wave[max_index]
    # print(max_refl)
    # print(max_index)
    # print(max_wave)

    dip_index = list_wave.index(dip)
    Rb = list_refl[dip_index]
    Tb = list_trans[dip_index]
    return max_refl, max_wave, Rb, Tb

def get_T_Max(sheet,dip):
    list_refl = []
    list_wave = [] 
    list_trans = []
    col_wave = sheet["K"]
    col_refl = sheet["L"]
    col_trans = sheet["M"]
    for i in range (4, len(col_refl)):
        if col_refl[i].value  != None:
            list_refl.append(col_refl[i].value)
        else:
            break
    for j in range (4, len(col_wave)):
        if col_wave[j].value  != None:
            list_wave.append(col_wave[j].value)
        else:
            break
    for k in range (4, len(col_trans)):
        if col_trans[k].value  != None:
            list_trans.append(col_trans[k].value)
        else:
            break   
    max_refl = max(list_refl)
    
    index_list = []
    for z in range(len(list_refl)):
        if list_refl[z] == max_refl:
            index_list.append(z)
    max_index = index_list[0]
    max_wave = list_wave[max_index]
   
    dip_index = list_wave.index(dip)
    Rt = list_refl[dip_index]
    Tt = list_trans[dip_index]
    return max_refl, max_wave, Rt, Tt

def get_F_byQW(sheet):
    notes_Col = sheet["G"]
    thick_Col = sheet["B"]
    mylist = []
    list_thick = []
    notes_list = []
    total = 0
    total_F = 0
    total_F_Up = 0
    total_F_Down = 0

    for n in range(3,len(notes_Col)):
        if notes_Col[n].value != None:
            notes_list.append(notes_Col[n].value)


    for x in range(len(notes_list)):
        if "{QW}" in notes_list[x]:
            mylist.append(x)
    
    for y in thick_Col:
        if y.value != None:
            list_thick.append(y.value)

    
    for i in range(3,len(list_thick)):
        total_F = total_F + list_thick[i]

    for j in range(mylist[-1]+1,len(list_thick)):
        total_F_Up = total_F_Up + list_thick[j]

    for k in range(3,mylist[0]):
        total_F_Down = total_F_Down + list_thick[k]

    for z in mylist:
        # print(list_thick[z])
        total += list_thick[z+2]
        
    # repeat_cell = sheet["D%d"%(mylist[0]+1)]
    total_QW_nm= total
    
    total_QW_cm= total_QW_nm*0.0000001
    return total_QW_nm, total_QW_cm, total_F, total_F_Up, total_F_Down


if __name__ == "__main__":
    a = True
    while a:
        print("-------------------------------------------------------------------------------")
        print("This program will generate simulation summary")
        print("-------------------------------------------------------------------------------")
        fname = str(input("Please input the exact name of the file you want to process: "))
        try:
            f = open(fname)
            a = False
        except FileNotFoundError:
            print('File does not exist')

    print(fname+" processing...")
    workbook = load_workbook(filename = fname)
    
    simul_sheet = workbook.active
    simul_sheet = workbook["SimulationData"]
    profile_sheet = workbook["Profile"]
    data_sheet = workbook["Design Inputs"]
    

    wave_F_dip = get_F_Dip(simul_sheet)
    B_max_refl, B_max_wave, Rb, Tb = get_B_Max(simul_sheet,wave_F_dip)
    T_max_refl, T_max_wave, Rt, Tt = get_T_Max(simul_sheet,wave_F_dip)

    

    new_list, new_list_A, new_list_B, new_list_C, new_list_n = composition_profile(data_sheet)

    X_list = get_X(new_list)
    print_profile_X(profile_sheet,X_list,fname)
    print_profile_thick(profile_sheet, new_list,fname)
    print_profile_notes(profile_sheet, new_list_n, fname)
    print_profile_A(profile_sheet, new_list_A,fname)
    print_profile_B(profile_sheet, new_list_B,fname)
    print_profile_C(profile_sheet, new_list_C,fname)


    workbook.save(filename = fname)
    qw_total_nm, qw_total_cm, total_F, total_Surf, total_Sub= get_F_byQW(profile_sheet)
    
    slope = (1239/wave_F_dip)*(Tt/(200-Rb-Rt))
    threshold = (1/(4*qw_total_cm))*(math.log(10000/(Rb*Rt)))
    
    print_E_profile(simul_sheet, profile_sheet)

    A_profile_x, A_Profile_a = A_Profile(profile_sheet)
    print_A_profile(profile_sheet, A_profile_x, A_Profile_a)

    B_profile_x, B_Profile_a = B_Profile(profile_sheet)
    print_B_profile(profile_sheet, B_profile_x, B_Profile_a)

    C_profile_x, C_Profile_a = C_Profile(profile_sheet)
    print_C_profile(profile_sheet, A_profile_x, C_profile_x, C_Profile_a)

    workbook.save(filename = fname)
    cell_dip = simul_sheet["AB6"]
    cell_dip.value = wave_F_dip

    cell_B_max_w = simul_sheet["AB7"]
    cell_B_max_w.value = B_max_wave

    cell_T_max_w = simul_sheet["AB8"]
    cell_T_max_w.value = T_max_wave

    cell_B_max_r = simul_sheet["AB12"]
    cell_B_max_r.value = B_max_refl

    cell_T_max_r = simul_sheet["AB13"]
    cell_T_max_r.value = T_max_refl

    cell_slope = simul_sheet["AB18"]
    cell_slope.value = slope

    cell_threshold = simul_sheet["AB19"]
    cell_threshold.value = threshold

    cell_qw_cm = simul_sheet["AB20"]
    cell_qw_cm.value = qw_total_cm

    cell_Rb= simul_sheet["AB24"]
    cell_Rb.value = Rb

    cell_Rt= simul_sheet["AB25"]
    cell_Rt.value = Rt

    cell_Tb= simul_sheet["AB26"]
    cell_Tb.value = Tb

    cell_Tt= simul_sheet["AB27"]
    cell_Tt.value = Tt

    cell_f= simul_sheet["AB32"]
    cell_f.value = total_F

    cell_f_up= simul_sheet["AB33"]
    cell_f_up.value = total_Surf

    cell_f_down= simul_sheet["AB34"]
    cell_f_down.value = total_Sub

    cell_qw_nm= simul_sheet["AB37"]
    cell_qw_nm.value = qw_total_nm

    workbook.save(filename = fname)

    print("WORK DONE")