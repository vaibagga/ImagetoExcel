import cv2
import xlsxwriter
from colormap import rgb2hex

# Create a workbook and add a worksheet.




class ImageProcessing:
    def __init__(self, image_path):
        self.image = cv2.imread(image_path, cv2.IMREAD_UNCHANGED)

    def toGrid(self, output_size):
        resized_image = cv2.resize(self.image, (100, 100), cv2.INTER_CUBIC)
        #cv2.imshow("image",resized_image)
        #cv2.waitKey(0)  # waits until a key is pressed
        #cv2.destroyAllWindows()
        return resized_image


class ExcelProcessing:
    def __init__(self):
        pass

    def getCol(self, n):
        result = ''

        while n > 0:
            index = (n - 1) % 26
            result += chr(index + ord('A'))
            n = (n - 1) // 26
        return result[::-1]

    def createExcel(self, image, output_path):
        rows, cols, _ = image.shape
        self.workbook = xlsxwriter.Workbook(output_path)
        self.worksheet = self.workbook.add_worksheet()
        for row in range(rows):
            for col in range(cols):
                #print(image[row][col])
                b, g, r =(image[row, col])
                cell_format = self.workbook.add_format()
                cell_format.set_pattern(1)  # This is optional when using a solid fill.
                cell_format.set_bg_color(f'{rgb2hex(r,g,b)}')
                excel_col = self.getCol(col+1)
                excel_cell_name = str(excel_col) + str(row+1)
                #print(rgb2hex(r,g,b))
                #print(excel_cell_name)
                self.worksheet.write(excel_cell_name, '', cell_format)
        self.workbook.close()

