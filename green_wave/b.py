inp = [1,2,3,1,2,3]
difs = []

for i in range(1, len(inp)):
    difs.append(inp[i]-inp[i-1])

if difs[0] > 0:
    plus = True
else:
    plus = False

ans = 1
cur_len = 1
for elem in difs:
    if (elem > 0 and plus) or (elem < 0 and not plus):
        cur_len += 1
    else:
        ans = max(ans, cur_len)
        cur_len = 1
        plus = not plus
