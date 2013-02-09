import os
import os.path
import sys
import shutil
import win32com.client

ppSaveAsShow = 7
ppSaveAsPNG = 18
ppSaveAsPDF = 32
ppLayoutBlank = 12


if len(sys.argv) < 2:
    print "Usage: ppt_rasterize.py Presentation.pptx"
    sys.exit(1)

pfilename = os.path.abspath(sys.argv[1])
path, filename = os.path.split(pfilename)
name = '.'.join(filename.split('.')[:-1])
slidesPath = os.path.join(path, 'PhotoAlbumSlides')
slideShowFileName = os.path.join(path, name + '.pps')
pdfFileName = os.path.join(path, name + '.pdf')
imagePdfFileName = os.path.join(path, name + '.img.pdf')

transitionMembers = [ 'AdvanceOnClick', 'AdvanceOnTime', 'AdvanceTime', 'Duration', 'EntryEffect', 'Hidden', 'Speed' ]

slide_size = None
Application = win32com.client.Dispatch("PowerPoint.Application")
try:
    Presentation = Application.Presentations.Open(pfilename)
    try:
        # save slide size
        slide_size = Presentation.PageSetup.SlideSize
        # save transitions
        transitions = []
        for slide in Presentation.Slides:
            d = {}
            for member in transitionMembers:
                if hasattr(slide.SlideShowTransition, member):
                    d[member] = getattr(slide.SlideShowTransition, member)
            transitions.append(d)
        # save slide images
        Presentation.SaveAs(slidesPath, ppSaveAsPNG, 0)
    finally:
        Presentation.Close()
    
    # create photo album
    try:
        photoAlbum = Application.Presentations.Add(True)
        # restore slide size
        if slide_size:
            photoAlbum.PageSetup.SlideSize = slide_size
        # restore slide images amd transitions
        slides = [x for x in os.listdir(slidesPath) if x.endswith('.PNG')]
        slides.sort(key=lambda x: int(x[5:-4]) )
        assert len(slides) == len(transitions)
        for i, fn in enumerate(slides):
            slide = photoAlbum.Slides.Add(photoAlbum.Slides.Count + 1, ppLayoutBlank)
            slide.Shapes.AddPicture(os.path.join(slidesPath, fn), False, True, 0, 0)
            for member in transitionMembers:
                if member in transitions[i]:
                    setattr(slide.SlideShowTransition, member, transitions[i][member])
        # save as PPS
        photoAlbum.SaveAs(slideShowFileName, ppSaveAsShow, 0)
        # save as PDF
        photoAlbum.SaveAs(imagePdfFileName, ppSaveAsPDF, 0)
        photoAlbum.Close()
    finally:
        # clean
        shutil.rmtree(slidesPath)
finally:
    Application.Quit()
