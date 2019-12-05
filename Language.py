import win32api
import win32con
import win32file
import os

# class Language_Class:
#      def __init__(self):
#          self.LanguageID_Culture = ""
#          self.LanguageID_Decimal = ""
#          self.LanguageID_MSIPath = ""
#          self.LanguageID_MSTPath = ""
#
# Language_En_US= Language_Class()
# Language_zh_CN= Language_Class()
# Language_zh_TW= Language_Class()

m_MSI_path = None
m_MSI_name = "VTM300_X64.msi"
m_wilangid_vbs_name = "wilangid.vbs"
m_wisubstg_vbs_name = "wisubstg.vbs"
m_Release_path = None
m_APP_path = None
Language_dict = {"zh-CN": 2052, "zh-TW": 1028, "en-US": 1033}  # 语言字典
ReleaseDirList = None


def CreateMST(Language_Culture_A, Language_Culture_B, ReleasePath, MSI_name):
    szTmpFile = None
    szTmpFile = os.path.join(ReleasePath, Language_Culture_A, MSI_name)
    if not os.path.exists(szTmpFile):
        win32api.MessageBox(0, "必要文件缺失！\n" + szTmpFile, "错误！", win32con.MB_OK | win32con.MB_ICONWARNING)
        exit()
    Language_MSI_A = szTmpFile

    szTmpFile = None
    szTmpFile = os.path.join(ReleasePath, Language_Culture_B, MSI_name)
    if not os.path.exists(szTmpFile):
        win32api.MessageBox(0, "必要文件缺失！\n" + szTmpFile, "错误！", win32con.MB_OK | win32con.MB_ICONWARNING)
        exit()
    Language_MSI_B = szTmpFile

    Language_MST_B = os.path.join(ReleasePath, Language_Culture_B, Language_Culture_B + ".mst")
    szCml = "-t language " + "\"" + Language_MSI_A + "\" \"" + Language_MSI_B + "\" -out  \"" + Language_MST_B + "\""
    win32api.ShellExecute(None, "open", "torch.exe", szCml, m_APP_path, 1)


if __name__ == "__main__":
    ##############################
    if m_MSI_name == None:
        win32api.MessageBox(0, "请先在PYthon脚本中配置编译生成后的 MSI 文件名！", "错误！",
                            win32con.MB_OK | win32con.MB_ICONWARNING)
        exit()
    #############################

    m_APP_path = os.getcwd()
    m_Release_path = os.path.join(os.getcwd(), "Release")

    # 检索Release文件夹中生成语言数
    ReleaseDirList = os.listdir(m_Release_path)
    n = ReleaseDirList.__len__()
    i = 0
    while i < n:
        str = ReleaseDirList[i]
        if Language_dict.get(str) != None:
            if str == "en-US":
                CreateMST("zh-CN", "en-US", m_Release_path, m_MSI_name)
            else:
                CreateMST("en-US", str, m_Release_path, m_MSI_name)
        i += 1

    #生成变形文件

