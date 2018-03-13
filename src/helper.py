import os
import shutil

def getMonthFolder(parentDir,month,monthIndex):
    path = parentDir + '\\' + str(monthIndex) +'_' + month
    return path


def copy_rename(src_file, dst_file):
        
        shutil.copyfile(src_file,dst_file)

monthNames = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']