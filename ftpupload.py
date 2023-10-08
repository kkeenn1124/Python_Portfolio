from ftplib import FTP
import operator,logging,openpyxl,datetime,shutil,os
 
 
def ftpconnect(host, username, password):
    ftp = FTP()  # 设置变量
    timeout = 30
    port = 21
    ftp.connect(host, port, timeout)  # 连接FTP服务器
    ftp.login(username,password)  # 登录
    return ftp
 
def downloadfile(ftp, remotepath, localpath,ffname):
    ftp.cwd(remotepath)  # 设置FTP远程目录(路径)
    list = ftp.nlst()  # 获取目录下的文件,获得目录列表
    print('檔案尋找中.........')
    for name in list:
        if(operator.contains(ffname,name) | operator.contains(ffname,'DM2320')):
            if(operator.contains(ffname,'DM2320')):
               name='SM2320'
            ftp.cwd(remotepath+'/'+name)
            filepath=ftp.nlst()
            for filelist in filepath:
                if(operator.contains(filelist,ffname)):
                    llen=len(filelist)-3
                    zz=filelist[llen:len(filelist)]
                    if(zz=='rar'):
                        exten='.rar'
                    elif(zz=='.7z'):
                        exten='.7z'
                    
                    fffname=ffname+exten
                    print(f"檔案將下載至{localpath}")
                    print('正在下載'+fffname+'..........')
                    path = localpath +'\\'+ fffname  # 定义文件保存路径
                    f = open(path, 'wb')  # 打开要保存文件
                    filename = 'RETR ' + fffname  # 保存FTP文件
                    ftp.retrbinary(filename, f.write)  # 保存FTP上的文件
                    print(fffname+' 下載完成')
                    break
            break
    ftp.set_debuglevel(0)         #关闭调试
    f.close()                    #关闭文件

def upload_file(ftp, local_path, remote_path,ffname):
    """
    # 从本地上传文件到ftp
    :param ftp:
    :param local_path:
    :param remote_path:
    :return:
    """
    try:
        ftp.cwd(remote_path)
        buff_size = 1024 * 1024 * 1024
        fp = open(local_path+ffname, 'rb')
        ftp.storbinary('STOR '+ffname, fp, buff_size)
        fp.close()
        print(newfile,'上傳成功')
    except Exception as e:
        logging.error('发送文件error:{}'.format(e))

##############################################################################


# filename=input('輸入程式檔名：')
dd=datetime.datetime.now().strftime('%Y-%m-%d')
path='.\\'
newfile=f'協力廠-華泰防毒資訊(COS001)(EMS){dd}.xlsx'

report='report'
if not os.path.isdir(report):
    os.mkdir(report)
mixfile=f'.\{report}\{newfile}'
shutil.copyfile(f'{path}協力廠-華泰防毒資訊(COS001)(EMS).xlsx',mixfile)
print('檔案',newfile,'建立成功')

wb=openpyxl.load_workbook(mixfile)
ws=wb.get_sheet_by_name('工作表1')
tt=datetime.datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')
dd=datetime.datetime.now().strftime('%Y年%m月%d日')
for i in range(ws.max_row-1):
    ws.cell(i+2,ws.max_column).value=dd
    ws.cell(i+2,ws.max_column-1).value=tt
wb.save(mixfile)



if __name__ == "__main__":
    try:
        ftp = ftpconnect('ftp.phison.com', "COS001B",'@!cos001')
        upload_file(ftp,path,'./Antivirus software information',newfile)
        ftp.quit()
    except:
        print('錯誤!!  請確認檔名是否有誤')


# input()