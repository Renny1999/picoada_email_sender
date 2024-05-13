f = open('skip.txt', 'r') 
for l in f.readlines():
  x = l.split(',')
  print(x[2])
