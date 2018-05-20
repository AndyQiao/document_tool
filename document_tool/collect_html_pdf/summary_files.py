import os

rootdif = 'reports'
lists = os.listdir(rootdif)

for path in lists:
    files = os.listdir('reports\\'+path)
    for file in files:
        sourceFile =  '.\\reports\\'+ path + '\\' + file
        targetFile = '.\\reports_summary\\' + file
        #print('source:', sourceFile)
        #print('targetFile:', targetFile)
        cmd = 'copy ' + sourceFile + ' ' + targetFile
        print(cmd)
        os.system ("copy %s %s" % (sourceFile, targetFile))