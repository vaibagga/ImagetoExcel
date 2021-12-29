import sys
from utils import *

def main():
    INPUT_IMAGE = sys.argv[1]
    OUTPUT_WIDTH = sys.argv[2]
    OUTPUT_HEIGHT = sys.argv[3]
    OUTPUT_SHEET = sys.argv[4]
    image_processor = ImageProcessing(INPUT_IMAGE)
    output_image = image_processor.toGrid((OUTPUT_WIDTH, OUTPUT_HEIGHT))
    excel_processor = ExcelProcessing()
    excel_processor.createExcel(output_image, OUTPUT_SHEET)






# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    main()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
