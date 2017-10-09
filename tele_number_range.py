import re
import os
import pickle
os.chdir(r"D:/ll")
file = open('ningbo.txt')


NingBoTelRange = []
for line in file:
    number = re.findall(r'\d{7}(?=<\/a><\/li>)',line)
    for i in number:
        NingBoTelRange.append(i)
print NingBoTelRange
file.close()
filew = open('ningbotel.txt','wb')
pickle.dump(NingBoTelRange,filew)
filew.close()