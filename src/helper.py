import os
import shutil

def getMonthFolder(parentDir,month):
    path = parentDir + '\\' + month
    return path


def copy_rename(sourceFolder, destFolder, old_file_name, new_file_name):
        src_file = os.path.join(sourceFolder, old_file_name)
        dst_file = os.path.join(destFolder, new_file_name)
        shutil.copyfile(src_file,dst_file)

monthNames = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']