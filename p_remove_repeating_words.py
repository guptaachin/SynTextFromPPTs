import os
d = set()
with open(os.path.join(os.getcwd(), 'data/new_words.txt'), 'r') as f:
    for line in f:
        line = line.rstrip()
        if line not in d:
            d.add(line)
        else:
            print(line)


if(len(d) > 0):
    f = open(os.path.join(os.getcwd(), 'data/new_words.txt'), 'w')
    for el in d:
        f.write(el+'\n')
    f.close()

print(len(d))

