"""
Inkscape use SimpInkScr for automation.
This file should be put under the extensions path in your local machine.
"""

# 置入白色底圖
rect((0, 84.2427),
     (340, 604),
     fill='white',
     stroke_width=0)  # 這是背景,(0, 84.2427),(344.032, 604.724)


# 置入灰色底圖
rect((0, 0),
     (340, 83.14),
     fill='gray',
     stroke_width=0)  # 這是背景,(0, 84.2427),(344.032, 604.724)

# 置入lulumi圖片
image_position = (41.5761, 136.5)
scale_size = 4.06  # 放大4.06倍
image("/Volumes/Seagate Bac/pyAutoGUI-demo-demo/inkscape-logo.png",
      image_position,
      transform=
      f'translate(-{image_position[0] * (scale_size - 1)}, -{image_position[1] * (scale_size - 1)}) '
      f'scale({scale_size}, {scale_size}) ',
      embed=True)  # 這是圖片, (39.8208, 134.049)

# 置入頁碼
t = text('1',
         (291.50,73.08),
         font_family='"Kozuka Mincho Pro", cursive',
         font_size='16pt',
         text_anchor='start',
         letter_spacing='3.46px'
         )

# 置入天數
t = text('20',
         (75, 518),
         font_family='"Mongolian Baiti", cursive',
         font_size='41pt',
         text_anchor='middle'
         )

# 置入月份
t = text('十月',
         (58.95, 456.23),
         font_family='"Kozuka Mincho Pro", cursive',
         font_weight='bold',
         font_size='10pt',
         text_anchor='start',
         letter_spacing='3.34px'
         )

# 置入星期
t = text('星期一',
         (44.48, 553.65),
         font_family='"Kozuka Mincho Pro", cursive',
         font_weight='bold',
         font_size='13pt',
         text_anchor='start',
         letter_spacing='3.46px'
         )

# 125, 8, 行距18.5


# for r in all_shapes():
#      # if r.tag == 'path':
#      #      print(r.__dict__)
#           print(r, r._transform)

# t = text('Hello, ', (canvas.width/2, canvas.height/2), font_size='24pt', text_anchor='middle')
# t.add_text('Inkscape', font_weight='bold', fill='#800000')
# t.add_text('!!!')
