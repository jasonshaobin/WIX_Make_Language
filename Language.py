import os
import win32api
import win32con

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
m_wilangid_vbs_path = None
m_wisubstg_vbs_path = None
m_Language_en_US_MSI = None
m_Release_path = None
m_APP_path = None
Language_dict = {"zh-CN": 2052, "zh-TW": 1028, "en-US": 1033}  # 语言字典
ReleaseDirList = None


def CreateMST(Language_Culture_A, Language_Culture_B):
    szTmpFile = os.path.join(m_Release_path, Language_Culture_A, m_MSI_name)
    if not os.path.exists(szTmpFile):
        win32api.MessageBox(0, "必要文件缺失！\n" + szTmpFile, "错误！", win32con.MB_OK | win32con.MB_ICONWARNING)
        exit()
    Language_MSI_A = os.path.join("Release", Language_Culture_A, m_MSI_name)
    # ------------------------------------
    szTmpFile = None
    szTmpFile = os.path.join(m_Release_path, Language_Culture_B, m_MSI_name)
    if not os.path.exists(szTmpFile):
        win32api.MessageBox(0, "必要文件缺失！\n" + szTmpFile, "错误！", win32con.MB_OK | win32con.MB_ICONWARNING)
        exit()
    Language_MSI_B = os.path.join("Release", Language_Culture_B, m_MSI_name)
    # ------------------------------------
    Language_MST_B = os.path.join("Release", "Transforms", Language_Culture_B + ".mst")
    szCml = "-t language " + "\"" + Language_MSI_A + "\" \"" + Language_MSI_B + "\" -out  \"" + Language_MST_B + "\""
    win32api.ShellExecute(None, "open", "torch.exe", szCml, m_APP_path, 1)
    szTmp=os.path.join(m_APP_path, Language_MST_B)
    if not os.path.exists(szTmp):
        win32api.MessageBox(0, "MST文件创建失败！\n" + szTmp, "错误！", win32con.MB_OK | win32con.MB_ICONWARNING)
        exit()
        return False
    return True


# WiSubStg.vbs "en-us\DIAViewSetup.msi" "transforms\zh-cn.mst"  2052
def MergeMST(Language_Culture_en_US, Language_Culture_xx, Language_xx_decimal):
    global m_Language_en_US_MSI
    szTmpFile = None
    szTmpFile = os.path.join(m_Release_path, Language_Culture_en_US, m_MSI_name)
    if not os.path.exists(szTmpFile):
        win32api.MessageBox(0, "必要文件缺失！\n" + szTmpFile, "错误！", win32con.MB_OK | win32con.MB_ICONWARNING)
        exit()
    m_Language_en_US_MSI =  os.path.join("Release", Language_Culture_en_US, m_MSI_name)

    szTmpFile = None
    szTmpFile = os.path.join(m_Release_path, "Transforms", Language_Culture_xx + ".mst")
    if not os.path.exists(szTmpFile):
        win32api.MessageBox(0, "必要文件缺失！\n" + szTmpFile, "错误！", win32con.MB_OK | win32con.MB_ICONWARNING)
        exit()
    Language_xx_MST = os.path.join("Release", "Transforms", Language_Culture_xx + ".mst")

    m_wisubstg_vbs_name = os.path.join(m_APP_path, "WiSubStg.vbs")
    szCml = "\"" + m_Language_en_US_MSI + "\" \"" + Language_xx_MST + "\" " + str(Language_xx_decimal)
    p=os.system("WiSubStg.vbs" + " " + szCml)
    print(p)
    return True


if __name__ == "__main__":

    if m_MSI_name is None:
        win32api.MessageBox(0, "请先在Python脚本中配置编译生成后的 MSI 文件名！", "错误！",
                            win32con.MB_OK | win32con.MB_ICONWARNING)
        exit()

    m_APP_path = os.getcwd()
    m_Release_path = os.path.join(os.getcwd(), "Release")

    # 检索Release文件夹中生成语言数
    ReleaseDirList = os.listdir(m_Release_path)
    n = ReleaseDirList.__len__()
    for LanguageID_Culture in Language_dict:
        Language_decimal = Language_dict[LanguageID_Culture]
        if Language_decimal is not None:
            if LanguageID_Culture == "en-US":
                CreateMST("zh-CN", "en-US")
            else:
                CreateMST("en-US", LanguageID_Culture)


    # 生成变形文件
    # WiSubStg.vbs "en-us\DIAViewSetup.msi" "transforms\zh-cn.mst"  2052
    languageList = []
    for LanguageID_Culture in Language_dict:
        Language_decimal = Language_dict[LanguageID_Culture]
        tmp = str(Language_decimal)
        languageList.append(tmp)
        if Language_decimal is not None:
            if LanguageID_Culture == "en-US":
                pass
            else:
                Language_decimal = Language_dict[LanguageID_Culture]
                MergeMST("en-US", LanguageID_Culture, Language_decimal)

    # WiLangId.vbs "en-us\TestInstaller.msi"  Package  1033, 1028, 2052
    szTmpFile = os.path.join(m_APP_path, "WiLangId.vbs")
    if not os.path.exists(szTmpFile):
        win32api.MessageBox(0, "必要文件缺失！\n" + szTmpFile, "错误！", win32con.MB_OK | win32con.MB_ICONWARNING)
        exit()
    m_wilangid_vbs_path = szTmpFile
    strID = ",".join(languageList)
    szCml = "\"" + m_Language_en_US_MSI + "\" Package " + strID
    p=os.system("WiLangId.vbs" + " " + szCml)
    print(p)
