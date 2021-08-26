import os
import shutil

def moveContract(path,listOfContract):
    os.chdir(path + '\\' + '\CustomerContract')
    for aContract in listOfContract:
        storageEntry = os.listdir()
        setOfStorageEntry = set(storageEntry)
        if aContract.companyName not in setOfStorageEntry:
            os.mkdir(path + '\\' + 'CustomerContract' + '\\' + aContract.companyName)
            shutil.move(aContract.path,path + '\\' + 'CustomerContract' + '\\' + aContract.companyName)
        else:
            try:
                shutil.move(aContract.path,path + '\\' + 'CustomerContract' + '\\' + aContract.companyName)
            except shutil.Error:
                aContract.error = True
                aContract.errorMsg += '\n Contracts Summarized'

