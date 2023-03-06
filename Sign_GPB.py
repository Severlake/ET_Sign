import os
import time
import random
import fitz
from fitz import TextPage
from PIL import Image

start_time = time.time()

print('DAYLESFORD ROBOTICS INC \u00A9\n')
####################### FOLDERS #######################
New = os.getcwd() + '/1_New'
FILES = os.listdir(New)
Ready = os.getcwd() + '/2_Ready'

width = 595
height = 842

signature_path_Tugaleva = os.getcwd() + '/Signature/Tugaleva'
signature_list_Tugaleva = os.listdir(signature_path_Tugaleva)


def company_name(CN):
      page = CN.load_page(0)
      text_page = page.get_textpage()
      words = TextPage.extractWORDS(text_page)

      coordinate = ''
      for i in words:
            if '/Н.М.' in i:
                  coordinate = (i[0], i[1], i[2], i[3])
                  break

      for i in words:
            if '7727325766' in i[4]:
                  return 'DMA', coordinate
      return 'DM', coordinate

def monochrome(pg, rotate=False):
      pix = pg.get_pixmap(dpi=160)
      pix.save('page.png')
      img = Image.open('page.png')#.convert('LA')
      if rotate:
            img = img.transpose(Image.ROTATE_90)

      img.save('page.png')

def pre_processing(taska, rotate=False):
      src = fitz.Document(New + '/' + f'{taska}')
      new_doc = fitz.Document()
      for ipage in src :
            print(f'{taska} == ', ipage.bound())
            if rotate:
                  fmt = fitz.paper_rect("A4-L")
            else:
                  fmt = fitz.paper_rect("A4")
            page = new_doc.new_page(width=fmt.width, height=fmt.height)
            page.show_pdf_page(page.rect, src, ipage.number)

      src.close()
      filename = taska.split('.')[:-1]
      new_name = filename[0] + '_1.pdf'
      new_file = New + '/' + f'{new_name}'
      new_doc.save(new_file)

def final_processing(doc,page,rotate=False):
      # ################ DELETE AND CLEAN PAGE #########################
      doc.delete_page(page)
      new_page = doc.new_page(page, width, height)
      if rotate:
            new_page.set_rotation(90)
      ################ PNG TO PDF CONVERTION #########################
      png = fitz.open('page.png')  # open image as a document
      pdf = png.convert_to_pdf()  # make a 1-page PDF of it
      png_to_pdf = fitz.open("pdf", pdf)  # open created pdf
      #################### PAGE ROTATION #############################
      bound = fitz.Rect(0.0, 0.0, width, height)
      range_rotation = [-1, -0.9, -0.8, -0.7, -0.6]
      rotation = random.choice(range_rotation)
      new_page.show_pdf_page(bound, png_to_pdf, rotate=rotation)  # make a page rotation
      #print('NEW_PAGE == ', new_page.bound())

def Signa(company, filename, coordinate):
      stamp_path = os.getcwd() + f'/Company/{company}'
      stamp_list = os.listdir(stamp_path)

      doc = fitz.Document(New + '/' + filename)
      for page in range(len(doc)):
            pg = doc[page]
            #print(pg.bound())
            ########################## STAMP ##############################
            x = random.randrange(0, 40, 5)
            y = random.randrange(0, 10, 1)
            #x = y = 0
            #rect_stamp = fitz.Rect(40 + x, 610 + y, 165 + x, 740 + y)
            rect_stamp = fitz.Rect(
                  coordinate[0] - 110 + x,
                  coordinate[1] - 20 + y,
                  coordinate[0] + 15 + x,
                  coordinate[3] + 110 + y)

            stamp_choose = random.choice(stamp_list)
            while stamp_choose == '.DS_Store' or stamp_choose == 'Thumbs.db':
                  stamp_choose = random.choice(stamp_list)
            stamp = open(stamp_path + '/' + stamp_choose, 'rb').read()

            pg.insert_image(rect_stamp, stream=stamp, overlay=True)

            ########################## SIGNATURE ##########################
            x = random.randrange(0, 40, 5)
            #x = random.randrange(0, 20, 2)
            #x1 = x2 = 0

            sign = random.choice(signature_list_Tugaleva)
            while sign == '.DS_Store' or sign == 'Thumbs.db':
                  sign = random.choice(signature_list_Tugaleva)
            signature = open(signature_path_Tugaleva + '/' + sign, 'rb').read()
            #rect_signature_Tugaleva = fitz.Rect(50 + x2, 610, 120 + x2, 650)
            rect_signature_Tugaleva = fitz.Rect(
                  coordinate[0] - 100 + x,
                  coordinate[1] - 10,
                  coordinate[0] - 30 + x,
                  coordinate[3] + 20)


            pg.insert_image(rect_signature_Tugaleva, stream=signature, overlay=True)

            ####################### FINAL PROCESS #########################
            monochrome(pg)
            final_processing(doc, page)
            doc.save(Ready + '/' + filename)
            #print(f'--- {taska} page: {page + 1} done ---')
      print("--- %s seconds ---" % (time.time() - start_time))


def main():
      for filename in FILES:
            if filename == '.DS_Store':
                  continue
            else:
                  CN = fitz.Document(New + '/' + f'{filename}')  # just to determine DM/DMA
                  company, coordinate = company_name(CN)
                  CN.close()
                  Signa(company, filename, coordinate)


      #os.remove('page.png')
      input('ALL DOCUMENTS COMPLETED: Press Enter to exit')

if __name__ == '__main__':
    main()
