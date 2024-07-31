inp = [100, -1, 300, 400, -5]

pref = [0]
last_pref = 0

for elem in inp:
    last_pref += elem
    pref.append(last_pref)

mins = []
cur_min = 0

for elem in pref[:-1]:
    if cur_min >= elem:
        cur_min = elem

    mins.append(cur_min)

difs = []

for elem, m in zip(pref, mins):
    difs.append(elem - m)

print(max(difs))

