"""
Inkscape use SimpInkScr for automation.
This file should be put under the extensions path in your local machine.
"""

image("C:\\Users\\123\\Downloads\\inkscape-logo.png", (0, 0), embed=True)
image("C:\\Users\\123\\Downloads\\Vector-based_example.svg", (0, 0), embed=True)

t = text('Hello, ', (canvas.width/2, canvas.height/2), font_size='24pt', text_anchor='middle')
t.add_text('Inkscape', font_weight='bold', fill='#800000')
t.add_text('!!!')
