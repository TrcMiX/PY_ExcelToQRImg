from qrcode.main import QRCode
from qrcode import constants
import xlrd
from tqdm import tqdm
from PIL import Image
import os

exFilePath = ''

if __name__ == '__main__':
    exFilePath = input('\n拖入文件 -> 回车 -> 等吧,很快\n')


def compressImage(srcPath, dstPath):
    for filename in tqdm(os.listdir(srcPath), ncols=70):
        if not os.path.exists(dstPath):
            os.makedirs(dstPath)

        srcFile = os.path.join(srcPath, filename)
        dstFile = os.path.join(dstPath, filename)

        if os.path.isfile(srcFile):
            try:
                sImg = Image.open(srcFile)
                w, h = sImg.size
                dImg = sImg.resize((int(w / 2), int(h / 2)), Image.ANTIALIAS)
                dImg.save(dstFile)
                # print(dstFile + " 成功！")
            except Exception:
                print(dstFile + "失败！")
    print('压缩完成!输出目录: %s\n-------------------------------------' % dstPath)


def getInfo():
    data = xlrd.open_workbook(exFilePath)
    try:
        for sheetname in data.sheet_names():
            table = data.sheet_by_name(sheetname)
            data_cols = table.ncols
            print('\n-------------------------------------\n%s 开始处理' % sheetname)
            col_values = table.col_values(0)
            muti_code(sheetname, col_values)
        input("\n-------------------------------------\n完事了,任意键关闭")
    except Exception as e:
        print(e)


class ConsoleOutputRedirect:
    def __init__(self, fp):
        self.fp = fp

    def write(self, s):
        self.fp.write(s)

    def writelines(self, lines):
        self.fp.writelines(lines)

    def flush(self):
        self.fp.flush()


def muti_code(sheetname, col_values):
    global temp_stdout
    s = 0
    # stdout_redirect = ConsoleOutputRedirect(sys.stdout)
    for machCode in tqdm(col_values, ncols=70):  # , ncols=50
        if machCode:
            # stdout_redirect.fp = StringIO()
            # temp_stdout, sys.stdout = sys.stdout, stdout_redirect

            filename = int(machCode)

            save_path = './qr_png/{}/{}'.format(sheetname, str(filename))
            makdirs(sheetname)
            create_qrcode("tyh://coupon/exchange?dataJson=%s" % str(filename), save_path)
            # myqr.run(level='H', words="tyh://coupon/exchange?dataJson=" + filename, version=1,save_name=save_path)  # picture='D:/111/qr_png/tyh.png', colorized=True,
            s += 1
        else:
            print("执行失败！%s-------第%s条！" % (sheetname, str(s)))
            pass

        # sys.stdout = temp_stdout
    # print("生成成功！第%s列,共%s条！\n-------------------------------------\n开始压缩!" % (str(i + 1), str(s)))
    # compressImage('D:/111/qr_png/0/', 'D:/111/qr_png/优惠券二维码/')


def makdirs(sheetname):
    if not os.path.isdir('./qr_png/{}'.format(sheetname)):
        os.mkdir('./qr_png/{}'.format(sheetname))


def create_qrcode(s, qrcodename):
    qr = QRCode(
        version=1,  # 设置容错率为最高
        error_correction=constants.ERROR_CORRECT_H,  # 用于控制二维码的错误纠正程度
        box_size=8,  # 控制二维码中每个格子的像素数，默认为10
        border=1,  # 二维码四周留白，包含的格子数，默认为4
        # image_factory=None, 保存在模块根目录的image文件夹下
        # mask_pattern=None
    )
    qr.add_data(s)  # QRCode.add_data(data)函数添加数据
    qr.make(fit=True)  # QRCode.make(fit=True)函数生成图片

    img = qr.make_image()
    img = img.convert("RGBA")  # 二维码设为彩色
    logo = Image.open('D:/111/qr_png/tyh.png')  # 传gif生成的二维码也是没有动态效果的

    w, h = img.size
    logo_w, logo_h = logo.size
    factor = 3  # 默认logo最大设为图片的四分之一
    s_w = int(w / factor)
    s_h = int(h / factor)
    if logo_w > s_w or logo_h > s_h:
        logo_w = s_w
        logo_h = s_h

    logo = logo.resize((logo_w, logo_h), Image.ANTIALIAS)
    l_w = int((w - logo_w) / 2)
    l_h = int((h - logo_h) / 2)
    logo = logo.convert("RGBA")
    img.paste(logo, (l_w, l_h), logo)
    img.save(os.getcwd() + '/' + qrcodename + '.png', quality=100)


getInfo()
