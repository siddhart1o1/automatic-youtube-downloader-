import ClointFusion as cf 
import time
import os
from pytube import YouTube

#FUNCTION TO DOWNLOAD VIDEO 
def vidoeDownloader(yt,quality):
    hightesQuality = "144p";

    for i in yt.streams:
        
        if(i.resolution==None):
            break

        len1 = len(hightesQuality)
        len2 = len(i.resolution)

        #FINDING THE HIGHEST QUALITY OF VIDEO AVIALBLE IF THE QUALITY IS NOT GIVEN THIS WILL BE USED
        if(int((i.resolution)[:len2-1]) > int((hightesQuality)[:len1-1])):
            hightesQuality = i.resolution

        #IF THE RESULUTION IS FOUND THEN VIDEO WILL BE DOWNLOADED
        if(i.resolution == quality):
            yt.streams.filter(res=i.resolution).first().download()
            print("DOWNLOADED in :",i.resolution)
            return

    #IF QULAITY IN EXCEL IS SET TO HIGHT OR EMPTY OR NOT GIVEN THEN THE HIGHEST QUALITY WILL BE DOWNLOADED 
    #OR IF THE QUALITY IS NOT FOUND THEN ALSO THE HIGHEST QUALITY WILL BE DOWNLOADED
    yt.streams.filter(res=hightesQuality).first().download()
    print("DOWNLOADED in :",hightesQuality)
    return



#YOUTUBE-LINKS DATA PATH
path = os.path.join(os.getcwd(),"youtube-data.xlsx")

#NUMBER OF LINKS
count = cf.excel_get_row_column_count(excel_path=path, sheet_name='Sheet1', header=0)
print("total count of videos is :",count[0]-1)
i = 0
for i in range(count[0]-1):
    if(cf.excel_get_single_cell(excel_path=path, sheet_name='Sheet1', header=0, columnName='Status', cellNumber=i)=="done"):
        continue
    else:
        url = cf.excel_get_single_cell(excel_path=path, sheet_name='Sheet1', header=0, columnName='Links', cellNumber=i)
        try: 
            yt = YouTube(url)
        except:
            print('CONNECTION ERROR')
            continue

        #EXTRACTING THE QUALITY OF THE VIDEO FROM THE EXCEL FILE
        quality = cf.excel_get_single_cell(excel_path=path, sheet_name='Sheet1', header=0, columnName='Quality', cellNumber=i)
        
        vidoeDownloader(yt,quality)
        
        #SETTING DONE WHEN VIDEO IS DOWNLOADED
        cf.excel_set_single_cell(excel_path=path, sheet_name='Sheet1', header=0, columnName='Status', cellNumber=float(i),  setText='done')

print("----------Download Complete-------")

