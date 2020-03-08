import csv
import datetime


def get_chats():
    with open("Log_Of_Chats.csv", "r", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        return list(reader)


def get_chats_table():
    with open("Log_Of_Chats_Table.csv", "r", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        return list(reader)


def get_horizontal_names(chats):
    return chats[0][1:]


def get_vertical_names(chats):
    names = []
    for row in chats:
        if row[0] != "":
            names.append(row[0])
    return names


def get_invalid_dates(chats):
    invalid = []

    for row in chats[1:]:
        for item in row[1:]:
            item = item.strip()
            if item != "":
                if len(item) != 6:
                    invalid.append(item)
    return invalid


def get_missing_chat_dates(chats):
    def get_name_for_column(chats, key):
        return chats[0][key + 1]

    def get_column_idx_for_person(chats, person):
        for key, item in enumerate(chats[0][1:]):
            if item == person:
                return key + 1
        return None

    def get_other_date_for_pair(chats, person_a, person_b):
        for row in chats[1:]:
            if row[0] == person_b:
                column_idx = get_column_idx_for_person(chats, person_a)
                return row[column_idx]

    invalid = []

    for row in chats[1:]:
        for key, item in enumerate(row[1:]):
            item = item.strip()
            if item != "":
                person_a = row[0]
                person_b = get_name_for_column(chats, key)

                date_a = item
                date_b = get_other_date_for_pair(chats, person_a, person_b).strip()

                if date_a != date_b:
                    invalid.append({"person_a": person_a, "person_b": person_b, "date_a": date_a, "date_b": date_b, "key": key})
    return invalid


def check_for_talking_to_self(chats):
    invalid = []

    for key, row in enumerate(chats[1:]):
        person = row[0]
        if row[key + 1] != "":
            invalid.append(person)

    return invalid


def check_for_impossible_chats(chats):
    invalid = []

    # Check horizontally
    for row in chats[1:]:
        person = row[0]
        dates = []
        for key, item in enumerate(row[1:]):
            item = item.strip()
            if item != "":
                dates.append(item)

        dupes = list(set([x for x in dates if dates.count(x) > 1]))
        if len(dupes) > 0:
            invalid.append({"person": person, "dupes": dupes, "dir": "h", "dates": dates})

    # Check vertically
    for key, person in enumerate(chats[0][1:]):
        dates = []
        for row in chats[1:]:
            item = row[key + 1].strip()
            if item != "":
                dates.append(item)

        dupes = list(set([x for x in dates if dates.count(x) > 1]))
        if len(dupes) > 0:
            invalid.append({"person": person, "dupes": dupes, "dir": "v", "dates": dates})

    return invalid


def reshape_to_table(chats):
    def get_name_for_row(chats, row_idx):
        return chats[row_idx][0]

    table = {}

    for row_idx, row in enumerate(chats[1:]):
        for col_idx, date in enumerate(row[1:]):
            date = date.strip()
            if date != "":
                if date not in table:
                    table[date] = {}

                table[date][col_idx + 1] = get_name_for_row(chats, row_idx + 1)

    return table


def write_new_csv(chats, dates, table):
    def get_row(date, table, names):
        row_names = ["" for name in names]
        for key in list(table[date].keys()):
            row_names[key - 1] = table[date][key]
        return row_names

    names = get_horizontal_names(chats)
    header = [""] + names

    with open("Log_Of_Chats_Table.csv", "w") as f:
        writer = csv.writer(f, delimiter=",", quotechar="\"", quoting=csv.QUOTE_MINIMAL)
        writer.writerow(header)

        for date in dates:
            writer.writerow([date] + get_row(date, table, names))


chats = get_chats()

names1 = get_horizontal_names(chats)
names2 = get_vertical_names(chats)
# print(list(set(names1) - set(names2)))

# for key, name in enumerate(names2):
#     if name != names1[key]:
#         print(name, names1[key])

# invalid = get_missing_chat_dates(chats)
# for item in invalid:
#     print(item)

# invalid = check_for_talking_to_self(chats)
# for item in invalid:
#     print(item)

# invalid = check_for_impossible_chats(chats)
# for item in invalid:
#     print(item)

table = reshape_to_table(chats)
dates = sorted(list(table.keys()), key=lambda x: datetime.datetime.strptime(x, "%b-%y"))
# for date in table:
#     print(date, table[date])

# for date in dates:
#     print(table[date])
#     names = ["" for name in names1]
#     for key in list(table[date].keys()):
#         print(key, table[date][key])
#         names[key] = table[date][key]
#     print(date, names)

# write_new_csv(chats, dates, table)


def check_for_dupe_names(chats_table):
    def get_number_of_occurences(name, names):
        return len([i for i in names if i == name])

    invalid = []

    for row in chats_table[1:]:
        names = [i for i in row[1:] if i != ""]
        for name in names:
            occurences = get_number_of_occurences(name, names)
            if occurences > 1:
                invalid.append({"date": row[0], "name": name, "occurences": occurences})
    return invalid


def check_for_matching_conversation_partners(chats_table):
    def get_column_idx_for_person(chats_table, person):
        for key, item in enumerate(chats_table[0][1:]):
            if item == person:
                return key + 1
        return None

    invalid = []

    for row in chats_table[1:]:
        for idx, person in enumerate(row[1:]):
            spoke_to = chats_table[0][idx + 1]
            column_idx = get_column_idx_for_person(chats_table, spoke_to)

            if person != row[column_idx]:
                invalid.append({"date": row[0], "person": person})

    return invalid


def get_chats_for_person(chats_table, person):
    for key, item in enumerate(chats_table[0][1:]):
        if item == person:
            return [i[key + 1] for i in chats_table[1:] if i[key + 1] != ""]


chats_table = get_chats_table()

# print(check_for_dupe_names(chats_table))

# print(check_for_matching_conversation_partners(chats_table))

print(get_chats_for_person(chats_table, "Keith"))
