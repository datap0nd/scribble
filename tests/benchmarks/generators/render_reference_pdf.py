from pathlib import Path
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
base=Path(__file__).resolve().parents[1]/'generated/scribble-test-kit-v1/evaluator-only'
c=canvas.Canvas(str(base/'reference-deck.pdf'),pagesize=(960,540),invariant=1)
c.setTitle('Atlas June review reference deck');c.setAuthor('Scribble synthetic benchmark')
for i in range(1,7):
    c.drawImage(ImageReader(str(base/'reference-deck-slides'/f'slide-{i}.png')),0,0,width=960,height=540)
    c.showPage()
c.save()
print(base/'reference-deck.pdf')
