import sys, getopt
import dataWrite, names

def main(argv):
   month = ''
   employeeName = ''
   department = ''
   try:
      opts, args = getopt.getopt(argv,"hm:e:d:",["month=","employeeName=","department="])
   except getopt.GetoptError:
      print('datacopy.py -m month -e employeeName -d department')
      sys.exit(2)
   for opt, arg in opts:
      if opt == '-h':
         print('datacopy.py -m <month> -e <employeeName> -d <department>')
         sys.exit()
      elif opt in ("-m", "--month"):
         month = arg
      elif opt in ("-e", "--employeeName"):
         employeeName = arg
      elif opt in ("-d", "--department") :
         department = arg
   
   if month==None or int(month)>12 or int(month)<1:
         raise Exception("incorrect month value, should be 1-12")
   if employeeName==None or len(employeeName)==0:
      if department==None or len(department)==0:
            for depart in names.departmentName:
                  employeeNames=names.departmentName[depart]
                  for name in employeeNames:
                        dataWrite.writeFile(int(month), name, depart)
      else:
            employeeNames=names.departmentName[department]
            for name in employeeNames:
                  dataWrite.writeFile(int(month), name, department)
   else:
      dataWrite.writeFile(int(month), employeeName, department)

if __name__ == "__main__":
   main(sys.argv[1:])


