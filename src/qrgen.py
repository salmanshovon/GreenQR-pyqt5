
import qrcode
from PIL import Image


def qrwlogo(logopath,logosize,data,color,VersioN,errorCorrection,boxSize,borderSize):

    basewidth=int(logosize)
    logo = Image.open(logopath)
    # adjust image size
    wpercent = (basewidth/float(logo.size[0]))
    hsize = int((float(logo.size[1])*float(wpercent)))
    logo = logo.resize((basewidth, hsize), Image.ANTIALIAS)
    QRcode = qrcode.QRCode(
    version=VersioN,
    error_correction=errorCorrection,
    box_size=boxSize,
    border=borderSize,
    )

    
    

    # adding URL or text to QRcode
    QRcode.add_data(data)

    # generating QR code
    QRcode.make()

    # taking color name from user
    QRcolor = color

    # adding color to QR code
    QRimg = QRcode.make_image(
            fill_color=QRcolor, back_color="white").convert('RGB')

    # set size of QR code
    pos = ((QRimg.size[0] - logo.size[0]) // 2,
            (QRimg.size[1] - logo.size[1]) // 2)
    QRimg.paste(logo, pos)

    # save the QR code generated
    # QRimg.save(fname+'.png')
    return QRimg

def simpleqr(data,color,VersioN,errorCorrection,boxSize,borderSize):
    QRcode = qrcode.QRCode(
    version=VersioN,
    error_correction=errorCorrection,
    box_size=boxSize,
    border=borderSize,
    )

    
    

    # adding URL or text to QRcode
    QRcode.add_data(data)

    # generating QR code
    QRcode.make()

    # taking color name from user
    QRcolor = color

    # adding color to QR code
    QRimg = QRcode.make_image(
            fill_color=QRcolor, back_color="white").convert('RGB')
    
    return QRimg

    
