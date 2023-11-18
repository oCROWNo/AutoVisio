import win32com.client as win32
from alive_progress import alive_bar
import os

if __name__ == '__main__':
    # Get current path
    curr_path = os.getcwd()
    # Get the name list of all files and folders in the current path
    filename_list = os.listdir(r"%s\\" % curr_path)
    # find visio files
    for fimename in filename_list:
        if ".vsdx" in fimename:
            print("Step1: Get Visio File")
            # 默认RGB图片存储路径
            rgb_folder = r"%s\%s_图片" %(curr_path,fimename.split(".")[0])
            # 若不存在文件夹，则创建文件夹
            if not os.path.exists(rgb_folder):
                os.mkdir(rgb_folder)
                print("Step2: Create RGB Picture Folder (%s) " % rgb_folder)
            else:
                print("Step2: RGB Picture Folder exist (%s) " % rgb_folder)
            # 打开Visio程序
            # appVisio = win32.gencache.EnsureDispatch("Visio.Application")   # 窗口可视
            appVisio = win32.gencache.EnsureDispatch("Visio.InvisibleApp")   # 窗口不可视
            appVisio.Visible = False    # 设置不可视
            print("Step3: Open Visio App As Invisible")
            # 设置Visio参数："另存为"操作的相关参数
            appVisio.Settings.RasterExportQuality = 100    # 质量100%
            appVisio.Settings.RasterExportColorFormat = 4  # RGB color format(the default for JPG)
            """ 
            SetRasterExportResolution(resolution, Width, Height, resolutionUnits)函数对应于Visio另存为的分辨率设置，参数如下：
            — — resolution:0(screen resolution),1(printer resolution),2(source resolution),3(custom resolution)
            — — Width:The raster export resolution width. Must be greater than or equal to 1
            — — Height:The raster export resolution height. Must be greater than or equal to 1
            — — resolutionUnits:0(Pixels per inch),1(Pixels per centimeter)
            See At Url: https://learn.microsoft.com/en-us/office/vba/api/visio.applicationsettings
            """
            appVisio.Settings.SetRasterExportResolution(3, 300, 300, 0)  # RGB color format(the default for JPG)
            print("Step4: Complete Export Settings")
            # 读取Visio文件
            # vdoc = appVisio.Documents.Open(r"%s\%s" %(curr_path,fimename))    # Open函数在文件未打开时使用
            """
            OpenEx(FileName, Flags)函数可指定打开方式
            — — visOpenCopy: 0x01 以副本方式打开
            — — visOpenRO: 0x02 
            — — visOpenDocked: 0x04
            — — visOpenDontList: 0x08
            — — visOpenMinimized: 0x10 
            — — visOpenRW: 0x20
            — — visOpenHidden: 0x40 文件在隐藏窗口打开
            — — visOpenMacrosDisabled: 0x80
            — — visOpenNoWorkspace: 0x100
            See At Url: https://learn.microsoft.com/en-us/office/vba/api/visio.documents.openex
            """
            vdoc = appVisio.Documents.OpenEx(r"%s\%s" %(curr_path,fimename), 0X40 + 0x01)
            print("Step5: Open %s" %(fimename))
            # 创建进度条
            print("Step4: Export Pages as JPG File...")
            with alive_bar(vdoc.Pages.Count, force_tty = True) as bar:
                for page in vdoc.Pages:
                    # print(r"%s\%s.jpg" % (rgb_folder, page.Name))
                    """
                    Export(FileName)函数导出文件
                    See At Url: https://learn.microsoft.com/en-us/office/vba/api/Visio.Page.Export
                    """
                    page.Export(r"%s\%s.jpg" % (rgb_folder, page.Name))
                    bar()
            # quit
            print("Step6: Complete Export and Quit Visio")
            appVisio.Quit()

