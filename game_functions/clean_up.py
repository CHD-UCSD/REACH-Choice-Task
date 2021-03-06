import os
import glob
import pprint as pp
from dir_function import get_homepath


dir_hierarchy = 1
homepath = get_homepath(dir_hierarchy,go_home=False)

def delete_files(file_path,condition):
    os.chdir(file_path)
    current_files = glob.glob('*.xls')
    print 'processing {}...'.format(file_path)
    if condition:
        #remove every data without subjectID
        inputfile = glob.glob('*.xls') + glob.glob('')
        for f in inputfile:
            if f.split('_')[0]=='': 
                print 'removing files with no subjectID', f
                os.remove(f)
        for cond in condition:
            inputfile = glob.glob(cond)
            if inputfile:
                for f in inputfile: 
                    print 'removing',cond, f
                    os.remove(f)
                current_files = glob.glob('*.xls')
            if not inputfile: print 'found no match for', cond
        print 'files remaining:',len(current_files)
    else: print 'no condition given, do nothing.'
    return current_files

complete_datapath = homepath + 'data/complete_data/'
datapath = homepath + 'data/'


conditions = ['*conflicted copy*.xls','*conflicted copy*.log','*test*.xls','*test*.log']
complete_data = delete_files(complete_datapath,conditions)
other_data = delete_files(datapath,conditions)
