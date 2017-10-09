import pickle
import os
def tel_filter():
    os.chdir(r'd:/ll')
    file = open('ningbotel.txt')
    tel_Filter = pickle.load(file)
    return tel_Filter
#   print tel_Filter

