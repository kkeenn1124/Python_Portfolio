from ftplib import FTP
import operator
import datetime,os,re
 
 
def ftpconnect(host, username, password):
    ftp = FTP()  # 设置变量
    timeout = 30
    port = 21
    ftp.connect(host, port, timeout)  # 连接FTP服务器
    ftp.login(username,password)  # 登录
    return ftp
 
def downloadfile(ftp, remotepath, localpath,ffname):
    if not(os.path.exists(localpath)):
        os.mkdir(localpath)
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
                    print()
                    print('********************請將檔案解壓縮後 執行以下步驟********************\n')
                    print('1.進入設定更改序列號開頭03001\n2.確認PID/VID是否與OP03-N0526一致\n3.修改Log存在位置 命名為工單\n')
                    print('*********************************************************************')
                    os.startfile(localpath)
                    break
            break
    ftp.set_debuglevel(0)         #关闭调试
    f.close()                    #关闭文件


##############################################################################

print('********************Longsys程式下載工具********************')
filename=input('輸入程式檔名：')

if __name__ == "__main__":
    try:
        ftp = ftpconnect('ftp2.longsys.com', "ose",'8U722e6o')
        downloadfile(ftp,'/OSE/U盘/FW/',"D:\Longsys_FW",filename)
        ftp.quit()
    except:
        print('錯誤!!  請確認檔名是否有誤')

print()
input('按Enter關閉視窗......')