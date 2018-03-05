import sys, getopt
import dataWrite, names

def main(argv):
   month = ''
   employeeName = ''
   try:
      opts, args = getopt.getopt(argv,"hm:e:",["month=","employeeName="])
   except getopt.GetoptError:
      print('datacopy.py -m month -e employeeName')
      sys.exit(2)
   for opt, arg in opts:
      if opt == '-h':
         print('datacopy.py -m <month> -e <employeeName>')
         sys.exit()
      elif opt in ("-m", "--month"):
         month = arg
      elif opt in ("-e", "--employeeName"):
         employeeName = arg
   
   if month==None or int(month)>12 or int(month)<1:
         raise Exception("incorrect month value, should be 1-12")
   if employeeName==None or len(employeeName)==0:
      for name in names.names:
         dataWrite.writeFile(int(month),name)
   else:
      dataWrite.writeFile(int(month),employeeName)

if __name__ == "__main__":
   main(sys.argv[1:])


