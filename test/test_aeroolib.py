#!/usr/bin/python3
TEMPLATE_NAME = "/test.odt"
TEMPLATE_STYLE = "/test_styles.odt"
TEST_RESULT = "./test_result.odt"

from aeroolib.plugins.opendocument import Template, OOSerializer
from io import BytesIO
from os.path import dirname, abspath

folder = dirname(abspath(__file__))
#===============================================================================

with open(folder+TEMPLATE_NAME, 'rb') as fp:
    file_data = fp.read()
    
with open(folder+TEMPLATE_STYLE, 'rb') as fp:
    style_data = fp.read()

template_io = BytesIO()
template_io.write(file_data)

style_io = BytesIO()
style_io.write(style_data)

serializer = OOSerializer(template_io, oo_styles=style_io)

basic = Template(source=template_io, serializer=serializer)

args = {'title': 'Successful test of Aeroo Reports library (aeroolib)',
        'some_unusedva_lue':['x','y','z']
        }

data = basic.generate(**args).render().getvalue()
with open(TEST_RESULT,"wb") as fp:
    fp.write(data)

