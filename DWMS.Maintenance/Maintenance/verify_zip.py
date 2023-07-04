import os

def unzipped_files(path,content):
    bad=[]
    for i in content:
        zipped=os.listdir(path+'\\'+i)
        folders=next(os.walk(path+'\\'+i))[1]
        for j in zipped:
            if j not in folders:
                if('zip' not in j):
                    bad.append(i)
                    break    
    print('Those are the folder with files non zipped\n' + str(bad))
    return bad

def wrong_zipped(path,bad):
    wrong=[]
    for i in bad:
        zipped=os.listdir(path+'\\'+i)
        folders=next(os.walk(path+'\\'+i))[1]
        for j in zipped:
            if j not in folders:
                if('.out' in j or '.drop' in j or '.java' in j or '.sort' in j):
                    wrong.append(i)
                    break
    print('Those are the folder with wrong process\n' + str(wrong))
    return wrong
    
def call(path):
    content = next(os.walk(path))[1]
    try:
        content.remove('badrecords')
        content.remove('CABSOUT')
        bad=unzipped_files(path, content)
        wrong=wrong_zipped(path,bad)
    except:
        bad=unzipped_files(path, content)
        wrong=wrong_zipped(path,bad)



