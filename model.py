import configparser
import pandas as pd
import pulp
from openpyxl.reader.excel import load_workbook
import datetime
import cij_creator
from cij_creator import resource_path


def is_available(availability_df, lecturer_index, day, time_slot):
    # function needs to return 0 if not available, 1 otherwise
    filtered_availability_df = availability_df[(availability_df['id'] == str(lecturer_index)) &
                                               (availability_df['day'] == day) &
                                               (availability_df['From_index'] <= time_slot) &
                                               (availability_df['Until_index'] >= time_slot)]
    if filtered_availability_df.empty:
        return 1
    else:
        return 0


def main():
    config = configparser.ConfigParser()
    config.read('config.ini')

    cij_creator.main()

    cij_path = resource_path(f'data/{config.get("files", "cij_matrix")}')
    data_path = resource_path(f'data/{config.get("files", "data_file")}')
    dist_path = resource_path(f'data/{config.get("files", "distances")}')
    # Convert the DataFrame to a dictionary
    df = pd.read_csv(cij_path, index_col=0)  # Assuming the row headers are in the first column
    # Convert column and row index values to integers
    df.columns = df.columns.astype(int)
    df.index = df.index.astype(int)
    c = {(i, j): df.loc[j, i] for i in df.columns for j in df.index}

    lecturers = pd.read_excel(data_path, sheet_name="lecturer_data")
    lecturers.set_index('id', inplace=True)
    lecturers['position'] = lecturers['position'].apply(lambda x: x.split(',')) #to allow multiple roles per person
    lec_data = lecturers.to_dict(orient='index')

    organizations = pd.read_excel(data_path, sheet_name="org_data")
    organizations.set_index('id', inplace=True)
    for d in ["first_date", "second_date", "third_date"]:
        organizations[d].astype(int)
    for h in ["first_time", "second_time", "third_time"]:
        organizations[h] = [time.strftime('%H:%M') for time in organizations[h]]
    org_data = organizations.to_dict(orient='index')

    dist_data = pd.read_csv(dist_path, header=None, names=["from", "to", "meters", "seconds"])

    # input
    org_num = list(organizations.index)
    total_volunteers = list(lecturers.index)
    lec_num = [i for i in total_volunteers if 'Lecturer' in lecturers.loc[i]['position']]
    guide_num = [i for i in total_volunteers if 'Guide' in lecturers.loc[i]['position']]
    days = range(1, 32)  # 31 days in October
    slots = range(1, 42)  # 41 slots of 15 min in a day

    # parameters
    workshop_len = 2
    lec_len = 1
    current_year = datetime.datetime.now().year

    # decision variables
    indices_x = [(i, j, d, s) for i in org_num for j in total_volunteers for d in days for s in slots]
    indices_f = [(i, j, d) for i in org_num for j in total_volunteers for d in days]
    indices_z = [(j, d) for j in total_volunteers for d in days]
    indices_c = [(i, j) for i in org_num for j in total_volunteers]

    x = pulp.LpVariable.dicts(name="x", indices=indices_x, cat=pulp.LpInteger, lowBound=0, upBound=1)
    z = pulp.LpVariable.dicts(name="z", indices=indices_z, cat=pulp.LpInteger, lowBound=0, upBound=1)
    f = pulp.LpVariable.dicts(name="f", indices=indices_f, cat=pulp.LpContinuous, lowBound=0,
                              upBound=18)  # runs from 8:00 to 18:00

    avl_date = ["first_date", "second_date", "third_date"]
    avl_time = ["first_time", "second_time", "third_time"]
    date_time_tuples = list(zip(avl_date, avl_time))

    # converting slots to hours for a constraint
    num_slots = 41
    start_hour = 8
    slot_interval = 15
    time_slots = {}
    inv_time_slots = dict()
    hour = start_hour
    minute = 0
    for slot in range(1, num_slots + 1):
        if minute == 60:
            hour += 1
            minute = 0

        time_slots[slot] = f"{hour:02d}:{minute:02d}"
        inv_time_slots[f"{hour:02d}:{minute:02d}"] = slot
        minute += slot_interval

    # availability
    calendar_data_path = resource_path(f'data/{config.get("files", "calendar_file")}')
    workbook = load_workbook(calendar_data_path)
    sheet_names = workbook.sheetnames
    illegal_sheet_names = ['OCTOBER', 'Slots']
    availability_df = pd.DataFrame()
    for sheet_name in set(sheet_names) - set(illegal_sheet_names):

        lect_ava = pd.read_excel(calendar_data_path, sheet_name=sheet_name)
        sheet_id = sheet_name.split('-', 1)[0]
        sheet_name = sheet_name.split('-', 1)[1].strip()
        lect_ava['id'] = sheet_id
        lect_ava['Name'] = sheet_name
        try:
            lect_ava['day'] = lect_ava['Date'].dt.strftime('%Y-%m-%d').apply(lambda x: int(x.split('-')[2].lstrip('0')))
            availability_df = pd.concat([availability_df, lect_ava])
        except:
            print(f"no rows for {sheet_name}")

    availability_df['From'] = [time.strftime('%H:%M') for time in availability_df['From']]
    availability_df['Until'] = [time.strftime('%H:%M') for time in availability_df['Until']]
    availability_df['From_index'] = availability_df['From'].apply(lambda x: int(inv_time_slots[x]))
    availability_df['Until_index'] = availability_df['Until'].apply(lambda x: int(inv_time_slots[x]))

    # objective function

    model = pulp.LpProblem("Scheduling Model", pulp.LpMaximize)
    model += pulp.lpSum([c[(i, j)] * x[(i, j, d, s)] for (i, j, d, s) in indices_x])

    # constraints

    for i in org_num:
        model += (pulp.lpSum([x[(i, j, d, s)] for d in days for s in slots for j in total_volunteers]) <= 1,
                  f"constraint_one_slot_per_org_{i}")

    # Constraint (9.2)
    for i in org_num:
        for j in total_volunteers:
            for d in days:
                for s in slots:
                    hour = time_slots[s]

                    # because the excel include hours and not slots, we transform the slots into hours values
                    if (d, hour) not in [(org_data[i]["first_date"], org_data[i]["first_time"]),
                                         (org_data[i]["second_date"], org_data[i]["second_time"]),
                                         (org_data[i]["third_date"], org_data[i]["third_time"])]:
                        model += (
                            x[(i, j, d, s)] - 0 == 0,
                            "constraint_9.2_" + str(i) + "_" + str(j) + "_" + str(d) + "_" + str(s))

    # Constraint (1)- non-overlap constraint- one lecturer is scheduled to an organization
    for i in org_num:
        for d in days:
            for s in slots:
                model += (pulp.lpSum([x[(i, j, d, s)] for j in total_volunteers]) <= 1,
                          "constraint_1_" + str(i) + "_" + str(d) + "_" + str(s))

    # Constraint (2)- non-overlap constraint- one organization is scheduled to a lecturer
    for j in total_volunteers:
        for d in days:
            for s in slots:
                model += (pulp.lpSum([x[(i, j, d, s)] for i in org_num]) <= 1,
                          "constraint_2_" + str(j) + "_" + str(d) + "_" + str(s))

    # Constraint (3)- for a specific organization in a specific date- one lecturer will be scheduled
    for d in days:
        for i in org_num:
            if org_data[i]["is_workshop"] == 0:
                # [org_data["id"]==i] and org_data[org_data["is_workshop"]==0]:
                model += (pulp.lpSum([x[(i, j, d, s)] for j in lec_num for s in slots]) <= 1,
                          "constraint_3_" + str(d) + "_" + str(i))
            else:
                model += (pulp.lpSum([x[(i, j, d, s)] for j in guide_num for s in slots]) <= 1,
                          "constraint_4_" + str(d) + "_" + str(i))

    # Constraint (5) - One lecturer per organization per available date
    for i in org_num:
        available_days = {org_data[i]["first_date"], org_data[i]["second_date"], org_data[i]["third_date"]}
        available_slots = set([slot for slot, time in time_slots.items() if
                               time in [org_data[i]["first_time"], org_data[i]["second_time"],
                                        org_data[i]["third_time"]]])

        if org_data[i]["is_workshop"] == 0:
            model += (
                pulp.lpSum([x[(i, j, d, s)] for j in lec_num for d in available_days for s in available_slots]) == 1,
                "constraint_5_" + str(i) + "_" + str(d)
            )
        else:
            model += (
                pulp.lpSum([x[(i, j, d, s)] for j in guide_num for d in available_days for s in available_slots]) == 1,
                "constraint_6_" + str(i) + "_" + str(d)
            )

    # Constraint (7)
    for j in lec_num:
        for d in days:
            for i in org_num:
                if org_data[i]["is_workshop"] == 0:
                    for s in slots:
                        model += (x[(i, j, d, s)] <= 1 - pulp.lpSum(
                            [x[(k, j, d, t)] for t in range(s + 1, min(s + 6, 42)) for k in org_num if k != i]),
                                  "constraint_7_" + str(i) + "_" + str(j) + "_" + str(d) + "_" + str(s))

    # Constraint (8)
    for j in guide_num:
        for i in org_num:
            if org_data[i]["is_workshop"] == 1:
                for d in days:
                    for s in slots:
                        # model += (x[(i,j,d,s)] <= 1 - pulp.lpSum([x[(i,j,d,t)] for t in range(s+1, s+10)]), "constraint_6_" + str(j) + "_" + str(i) + "_" + str(d) + "_" + str(s))
                        model += (
                            x[(i, j, d, s)] <= 1 - pulp.lpSum([x[(i, j, d, t)] for t in range(s + 1, min(s + 10, 42))]),
                            f"constraint_8_{j}_{i}_{d}_{s}")

    # Constraint (9) - Availability
    for i in org_num:
        for j in total_volunteers:
            for d in days:
                for s in slots:
                    model += (x[(i, j, d, s)] <= is_available(availability_df, j, d, s),
                              "constraint_9_" + str(i) + "_" + str(j) + "_" + str(d) + "_" + str(s))

    # Constraint (9.1)
    for i in org_num:
        for j in total_volunteers:
            model += (
                pulp.lpSum(x[(i, j, d, s)] for d in days for s in slots) <= 1,
                "constraint_9.1_" + str(i) + "_" + str(j)
            )

    # Constraint (10)
    for j in total_volunteers:
        for d in days:
            model += (
                pulp.lpSum([x[(i, j, d, s)] for i in org_num for s in slots]) <= lec_data[j]["vol_limit"] + z[(j, d)],
                "constraint_10_" + str(j) + "_" + str(d))
    # Create a set of valid organization combinations to iterate over
    valid_combinations = [(k, i) for k in org_num for i in org_num if i != k]

    # Precompute distance values and availability for each organization
    distances = {}
    availability = {}
    for k, i in valid_combinations:
        dist = float(dist_data.loc[(dist_data['from'] == org_data[k]["address"]) & (
                    dist_data['to'] == org_data[i]["address"]), 'seconds'].values[0])
        distances[(k, i)] = dist
        availability[(k, i)] = (
            set([org_data[k]["first_date"], org_data[k]["second_date"], org_data[k]["third_date"]]),
            set([org_data[i]["first_date"], org_data[i]["second_date"], org_data[i]["third_date"]]),
            set([slot for slot, time in time_slots.items() if
                 time in [org_data[k]["first_time"], org_data[k]["second_time"], org_data[k]["third_time"]]]),
            set([slot for slot, time in time_slots.items() if
                 time in [org_data[i]["first_time"], org_data[i]["second_time"], org_data[i]["third_time"]]])
        )

    # Constraint (11)
    for k, i in valid_combinations:
        for j in total_volunteers:
            days_i, days_k, slots_i, slots_k = availability[(k, i)]
            for d in days_i & days_k:
                for s in slots_i & slots_k:
                    for o in range(1, s):
                        # Lecturer or guide
                        lec_length = lec_len if j in lec_num else workshop_len

                        dist = distances[(k, i)]

                        model += (
                            f[(i, j, d)] >= f[(k, j, d)] + lec_length + 0.5 * float(
                                (1 - lec_data[j]["mobility"])) * dist
                            + dist - 10000 * (2 - x[(k, j, d, o)] - x[(i, j, d, s)])
                            , "constraint_11_" + str(k) + "_" + str(i) + "_" + str(j) + "_" + str(d) + "_" + str(
                                s) + "_" + str(o)
                        )

    # Constraint (12)
    for i in org_num:
        for j in total_volunteers:
            for d in days:
                model += (f[(i, j, d)] == pulp.lpSum([x[(i, j, d, s)] * (8 + ((s - 1) * 15) / 60) for s in slots]),
                          "constraint_12_" + str(i) + "_" + str(j) + "_" + str(d))

    # running
    print("Solving Optimization Problem:")
    model.solve()

    # Check the status of the solution
    if model.status == pulp.LpStatusOptimal:
        # Retrieve and print the optimal objective value
        print("Total Cost = ", pulp.value(model.objective))

        # Retrieve and print the values of x[(i, j, d, s)]
        for item in indices_x:
            if (x[item].varValue != 0.0):
                print("x-" + str(item), x[item].varValue)

        # Retrieve and print the values of u[(i, j, d)]
        for item in indices_f:
            if (f[item].varValue != 0.0):
                print("f-" + str(item), f[item].varValue)

    else:
        print("Optimization problem did not find an optimal solution.")

    # ----------------------------------------------------------------------------------------------------------------------

    # Create a DataFrame to store the results
    results_df = pd.DataFrame(columns=['Organization', 'Volunteer', 'Date', 'Start Time', 'End Time', 'Location', 'Type'])


    # Iterate over the indices_x and retrieve the values
    for item in indices_x:
        if x[item].varValue != 0.0:
            org_id, volunteer_id, day, slot = item
            organization_name = org_data[org_id]['org_name']
            volunteer_name = lec_data[volunteer_id]['full_name']
            day_str = f"{day}/10/{current_year}"
            Start_time = time_slots[slot]
            results_df = pd.concat([results_df, pd.DataFrame({'Organization': [organization_name],
                                                              'Volunteer': [volunteer_name],
                                                              'Date': [day_str],
                                                              'Start Time': [Start_time],
                                                              'End Time':[time_slots[slot + (8 if org_data[org_id]['is_workshop']==1 else 4)]],
                                                              'Location':[org_data[org_id]['address']],
                                                              'Type': 'Workshop' if org_data[org_id]['is_workshop']==1 else 'Lecture'})])


    # Save the results to an Excel file
    print("Saving results to file")
    results_file = config.get('files', 'output_file')
    results_df.to_excel(f'{results_file}.xlsx', index=False)
    print("Finished Calculations: Results can be found in results.xlsx file")
    return results_df


if __name__ == '__main__':
    print("Starting calculation of best way to assign lecturers to organizations")
    try:
        main()
    except Exception as e:
        print("Error: ", e)
    finally:
        print("Finished Calculations: Results can be found in results.xlsx file")
        input("Press Enter to exit")
