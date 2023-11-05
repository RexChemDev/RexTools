import rextools as rt


data = rt.Worker("OT2_test_data.xlsx", 0, 1, "D", 4)

for com in data:
    print(com)
    print(data.is_finished)

print(len(data.commands))