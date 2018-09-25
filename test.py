import threading

pools = []
def a(st):
    try:
        print(st)
        if st == 1:
            raise Exception
    except Exception:
        print('error')



for i in range(5):
    t = threading.Thread(target=a, args=(i,))
    pools.append(t)


[t.start() for t in pools]
[t.join() for t in pools]
print(t)
