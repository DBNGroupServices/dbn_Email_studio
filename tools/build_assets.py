"""Create the app icon and disposable example emails used in local QA."""
from email.message import EmailMessage
from email import policy
from pathlib import Path
import io
import struct

from PIL import Image, ImageDraw, ImageFont

ROOT=Path(__file__).resolve().parents[1]
(ROOT/'assets').mkdir(exist_ok=True)
(ROOT/'examples').mkdir(exist_ok=True)
(ROOT/'web'/'icons').mkdir(exist_ok=True)
image=Image.new('RGBA',(1024,1024),(0,0,0,0))
draw=ImageDraw.Draw(image)
draw.rounded_rectangle((0,0,1023,1023),radius=256,fill='#4263eb')
draw.rounded_rectangle((192,292,828,748),radius=84,outline='#ffffff',width=40)
draw.line((212,344,512,548,804,344),fill='#ffffff',width=40,joint='curve')
draw.rounded_rectangle((672,132,928,388),radius=84,fill='#4263eb')
draw.polygon([(804,156),(828,232),(904,256),(828,280),(804,356),(780,280),(704,256),(780,232)],fill='#e3eaff')
image.resize((512,512),Image.Resampling.LANCZOS).save(ROOT/'assets'/'app-icon.png')
icon=image.resize((256,256),Image.Resampling.LANCZOS)
for name in ('app.ico','email-studio-blue.ico'):
    icon.save(ROOT/'assets'/name,sizes=[(16,16),(24,24),(32,32),(48,48),(64,64),(128,128),(256,256)])
icon.save(ROOT/'web'/'icons'/'email-blue-v2.ico',sizes=[(16,16),(24,24),(32,32),(48,48),(64,64),(128,128),(256,256)])
for size in (16,32,64,80,128,192,512):
    image.resize((size,size),Image.Resampling.LANCZOS).save(ROOT/'web'/'icons'/f'email-blue-v2-{size}.png')

message=EmailMessage()
message['From']='Alex <alex@example.com>'
message['To']='Email Studio <studio@example.com>'
message['Subject']='Security coverage for tenant event'
message['Date']='Tue, 27 Oct 2026 10:00:00 +1100'
message.set_content('Hi,\n\nCould you please check whether two security officers are available for our event on Friday, 30 October, from 5:00 pm to 9:00 pm? We expect approximately 80 guests.\n\nPlease let us know what information you need to prepare a quote.\n\nKind regards,\nAlex\n\nThis is a fictional example for testing Email Studio.')
message.add_attachment(b'Fictional test event. Guest count: 80. Venue to be confirmed.',maintype='text',subtype='plain',filename='Event notes.txt')
(ROOT/'examples'/'Example incoming email.eml').write_bytes(message.as_bytes(policy=policy.default))
(ROOT/'examples'/'Example dot points.txt').write_text('Thank Alex for the enquiry.\nWe will check officer availability and prepare a quote.\nAsk for the event location and special access requirements.\nDo not confirm the booking yet.\n',encoding='utf-8')

# A small, valid Outlook compound file, generated from fictional test data.
from extract_msg.ole_writer import OleWriter
writer=OleWriter()
writer.addEntry('__properties_version1.0',b'\x00'*32 + struct.pack('<IIII',0x340D0003,0,0x40000,0))
for tag,text in {'001A':'IPM.Note','0037':'Security coverage for tenant event','0C1A':'Alex','0C1F':'alex@example.com','0E04':'Email Studio','1000':message.get_body().get_content()}.items():
    writer.addEntry('__substg1.0_'+tag+'001F',text.encode('utf-16-le'))
buffer=io.BytesIO();writer.write(buffer)
(ROOT/'examples'/'Example Outlook email.msg').write_bytes(buffer.getvalue())
print('App icon and fictional examples created.')
