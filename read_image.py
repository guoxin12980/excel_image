from win32com.client.gencache import EnsureDispatch
from PIL import ImageGrab

ex = EnsureDispatch('Excel.Application')

wb = ex.Workbooks.Open(r'C:\Users\Gordon\PycharmProjects\testdistribute\test.xlsx')
num=10
for i in wb.Worksheets:
    for n,shape in enumerate(i.Shapes):
        shape.Copy()
        image= ImageGrab.grabclipboard()
        if image.size >= (10,10000000):
            image.convert('RGB').save('{}.png'.format(num),'jpeg')
        print(num)
        num+=1

ex.Quit()