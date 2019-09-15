#%%
import docx2txt
import os
from os.path import abspath
from googletrans import Translator
from difflib import SequenceMatcher
import nltk
from nltk.tokenize import sent_tokenize
import re
import shutil
from zipfile import ZipFile
base_path=os.path.dirname(abspath('__file__'))
#%%
def get_jc(path=base_path):
    jc=None
    if jc==None:
        for i in os.listdir(path):
            parts=i.split('_')
            if parts[-1].split('.')[-1]=='zip':
                jc=parts[0]+'_'+parts[1]
    return jc
job_code=get_jc()
#%%
def zip_to_txt(path=base_path):
    source=[]
    translated=[]
    source_names=[]
    for i in os.listdir(path):
        extension=i.split('.')[-1]
        if extension=='zip':
            if i.split('_')[2]=='Original':
                zf=ZipFile(i)
                zf.extractall()
                for j in os.listdir(path):
                    ext2=j.split('.')[-1]
                    if ext2==j:
                        docs_path=base_path+'\\'+j+'\\'+'Job_files'
                        for k in os.listdir(docs_path):
                            if k.split('.')[-1]=='docx':
                                source.append(docx2txt.process(docs_path+'\\'+k))
                                source_names.append(k)
                        shutil.rmtree(j)
            elif i.split('_')[2]=='Translate.zip':
                zf=ZipFile(i)
                zf.extractall()
                for j in os.listdir(path):
                    ext2=j.split('.')[-1]
                    if ext2==j:
                        docs_path=base_path+'\\'+j+'\\'+i.split('.')[0]
                        for k  in os.listdir(docs_path):
                            if k.split('.')[-1]=='docx':
                                translated.append(docx2txt.process(docs_path+'\\'+k))
                        shutil.rmtree(j)            
    source_str=''.join([text+'\n' for text in source])
    trans_str=''.join([text+'\n'for text in translated])
    return source_str,trans_str,source_names
#%%
source,translated,source_file_names=zip_to_txt()
# source=docx2txt.process('Source.docx')
# translated=docx2txt.process('Translate.docx')
# len(source),len(translated)
#%%
similarity=SequenceMatcher(None,source,translated)
rep=0
for i in similarity.get_matching_blocks():
    if i[2]>30:
        if source[i[0]-rep] in ['.',' ']:
                source=source.replace(source[i[0]+1-rep:i[0]+i[2]-rep],'',1)
                translated=translated.replace(translated[i[1]+1-rep:i[1]+i[2]-rep],'',1)
                rep += i[2]-1
        else:
                source=source.replace(source[i[0]-rep:i[0]+i[2]-rep],'',1)
                translated=translated.replace(translated[i[1]-rep:i[1]+i[2]-rep],'',1)
                rep += i[2]
#%%
languages={'en':'english','pt':'portuguese','ko':'korean','ja':'japanese'}

def doc_split(doc):
    translator=Translator()
    lan=translator.detect(doc[:1000]).lang
    if lan=='ko':
        tokens=doc.split('.')
        tokens=[i+'.' for i in tokens]
    elif lan=='ja':
        tokens=doc.split('。')
        tokens=[i+'。' for i in tokens]
    else:
        tokens=sent_tokenize(doc,language=languages[lan])
    split=[]
    len_counter=0
    temp_list=[]
    final=[]
    for i in range(len(tokens)):
        if len_counter+len(tokens[i])+len(temp_list)-1<2000:
            len_counter=len_counter+len(tokens[i])
            temp_list.append(tokens[i])
            #print(i,len_counter,len(temp_list))
        else:
            len_counter=len(tokens[i])
            split.append(temp_list)
            temp_list=[]
            temp_list.append(tokens[i])
            #print(i,len_counter,len(temp_list))
    split.append(temp_list)
    final=[''.join(i) for i in split]
    return final
#%%
google_output=[]
for i in doc_split(source):
    translator=Translator()
    google_output.append(translator.translate(i,dest='en').text)
google_translated=' '.join(google_output)
#%%
similarity2=SequenceMatcher(None,google_translated,translated)
matches=similarity2.get_matching_blocks()

ratio=similarity2.ratio()

print(ratio)

#%%
texts={'google_translated':google_translated,'translated':translated}
for i in texts.keys():
    fname=open('{}.txt'.format(i),'+w',encoding='utf8')
    fname.write(texts[i])
    fname.close()
#%%

results=open('all_matches.txt','a',encoding='utf8')

for i in matches:
    if i[2]>100:
        results.write(google_translated[i[0]:i[0]+i[2]]+'\n')

results.close()