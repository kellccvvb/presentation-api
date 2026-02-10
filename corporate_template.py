from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

class CorporatePresentation:
    """
    Corporate presentation template matching your brand guidelines.
    Colors: Deep Teal (#2C5F7C), Bright Red (#E31E24), Light Blue (#7BA7BC)
    Style: Modern, clean, split-screen layouts
    """
    
    def __init__(self):
        self.prs = Presentation()
        
        # Set slide dimensions (16:9 widescreen)
        self.prs.slide_width = Inches(10)
        self.prs.slide_height = Inches(7.5)
        
        # Brand colors
        self.TEAL = RGBColor(44, 95, 124)
        self.RED = RGBColor(227, 30, 36)
        self.LIGHT_BLUE = RGBColor(123, 167, 188)
        self.WHITE = RGBColor(255, 255, 255)
        self.LIGHT_GRAY = RGBColor(242, 242, 242)
        self.DARK_GRAY = RGBColor(51, 51, 51)
        
        self.FONT_NAME = "Calibri"
    
    def add_title_slide(self, title, subtitle):
        """Create title slide with teal panel"""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
        # Teal background panel (left side)
        teal_panel = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(0), Inches(1.5),
            Inches(6.5), Inches(5)
        )
        teal_panel.fill.solid()
        teal_panel.fill.fore_color.rgb = self.TEAL
        teal_panel.line.fill.background()
        
        # Red accent (top right)
        red_accent = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(9), Inches(0),
            Inches(1), Inches(1.5)
        )
        red_accent.fill.solid()
        red_accent.fill.fore_color.rgb = self.RED
        red_accent.line.fill.background()
        
        # Title text
        title_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(2.5),
            Inches(5.5), Inches(2)
        )
        tf = title_box.text_frame
        tf.text = title
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.font.name = self.FONT_NAME
        p.font.size = Pt(54)
        p.font.bold = True
        p.font.color.rgb = self.WHITE
        
        # Subtitle
        subtitle_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(4.8),
            Inches(5.5), Inches(1.2)
        )
        stf = subtitle_box.text_frame
        stf.text = subtitle
        stf.word_wrap = True
        
        sp = stf.paragraphs[0]
        sp.font.name = self.FONT_NAME
        sp.font.size = Pt(16)
        sp.font.color.rgb = self.WHITE
        
        return slide
    
    def add_table_of_contents(self, sections):
        """Create table of contents slide"""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
        # Red accent (top right)
        red_accent = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(9), Inches(0),
            Inches(1), Inches(1)
        )
        red_accent.fill.solid()
        red_accent.fill.fore_color.rgb = self.RED
        red_accent.line.fill.background()
        
        # Light blue accent
        blue_accent = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(9), Inches(1),
            Inches(1), Inches(0.8)
        )
        blue_accent.fill.solid()
        blue_accent.fill.fore_color.rgb = self.LIGHT_BLUE
        blue_accent.line.fill.background()
        
        # Title
        title_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(0.8),
            Inches(5), Inches(0.8)
        )
        tf = title_box.text_frame
        tf.text = "Table Of Content"
        
        p = tf.paragraphs[0]
        p.font.name = self.FONT_NAME
        p.font.size = Pt(40)
        p.font.bold = True
        p.font.color.rgb = self.DARK_GRAY
        
        # Accent line
        line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(0.5), Inches(1.5),
            Inches(1.2), Inches(0.03)
        )
        line.fill.solid()
        line.fill.fore_color.rgb = self.DARK_GRAY
        line.line.fill.background()
        
        # Add sections in two columns
        left_x = Inches(0.5)
        right_x = Inches(3.5)
        start_y = Inches(2.3)
        spacing = Inches(1.1)
        
        for i, section in enumerate(sections):
            if i % 2 == 0:
                x_pos = left_x
                y_pos = start_y + (i // 2) * spacing
            else:
                x_pos = right_x
                y_pos = start_y + (i // 2) * spacing
            
            # Number
            num_box = slide.shapes.add_textbox(x_pos, y_pos, Inches(0.8), Inches(0.5))
            ntf = num_box.text_frame
            ntf.text = f"{i+1:02d}."
            
            np = ntf.paragraphs[0]
            np.font.name = self.FONT_NAME
            np.font.size = Pt(32)
            np.font.bold = True
            np.font.color.rgb = self.DARK_GRAY
            
            # Section title
            text_box = slide.shapes.add_textbox(
                x_pos, y_pos + Inches(0.5),
                Inches(2.5), Inches(0.5)
            )
            ttf = text_box.text_frame
            ttf.text = section
            ttf.word_wrap = True
            
            tp = ttf.paragraphs[0]
            tp.font.name = self.FONT_NAME
            tp.font.size = Pt(14)
            tp.font.color.rgb = self.DARK_GRAY
        
        return slide
    
    def add_content_with_icons(self, title, items):
        """
        Create content slide with icon-text pairs
        items = [{"icon": "💡", "text": "Description"}, ...]
        """
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
        # Dark teal panel (right side)
        right_panel = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(4), Inches(0),
            Inches(6), Inches(7.5)
        )
        right_panel.fill.solid()
        right_panel.fill.fore_color.rgb = self.TEAL
        right_panel.line.fill.background()
        
        # Title (left side)
        title_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(1.2),
            Inches(3), Inches(2)
        )
        tf = title_box.text_frame
        tf.text = title
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.font.name = self.FONT_NAME
        p.font.size = Pt(36)
        p.font.bold = True
        p.font.color.rgb = self.DARK_GRAY
        
        # Accent line
        line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(0.5), Inches(3),
            Inches(1.2), Inches(0.03)
        )
        line.fill.solid()
        line.fill.fore_color.rgb = self.DARK_GRAY
        line.line.fill.background()
        
        # Add icon-text pairs
        start_y = Inches(1.5)
        spacing = Inches(1.9)
        
        for i, item in enumerate(items[:4]):
            y_pos = start_y + (i * spacing)
            
            # Red square for icon
            icon_bg = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(4.5), y_pos,
                Inches(0.7), Inches(0.7)
            )
            icon_bg.fill.solid()
            icon_bg.fill.fore_color.rgb = self.RED
            icon_bg.line.fill.background()
            
            # Icon
            icon_box = slide.shapes.add_textbox(
                Inches(4.5), y_pos,
                Inches(0.7), Inches(0.7)
            )
            itf = icon_box.text_frame
            itf.text = item.get("icon", "✓")
            itf.vertical_anchor = MSO_ANCHOR.MIDDLE
            
            ip = itf.paragraphs[0]
            ip.alignment = PP_ALIGN.CENTER
            ip.font.size = Pt(28)
            
            # Description
            text_box = slide.shapes.add_textbox(
                Inches(5.4), y_pos,
                Inches(4), Inches(1.5)
            )
            ttf = text_box.text_frame
            ttf.text = item["text"]
            ttf.word_wrap = True
            
            tp = ttf.paragraphs[0]
            tp.font.name = self.FONT_NAME
            tp.font.size = Pt(14)
            tp.font.color.rgb = self.WHITE
        
        return slide
    
    def add_split_slide(self, title, paragraphs):
        """Create 50/50 split slide with text"""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
        # Teal panel (right 50%)
        right_panel = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(5), Inches(0),
            Inches(5), Inches(7.5)
        )
        right_panel.fill.solid()
        right_panel.fill.fore_color.rgb = self.TEAL
        right_panel.line.fill.background()
        
        # Title
        title_box = slide.shapes.add_textbox(
            Inches(5.5), Inches(1.5),
            Inches(4), Inches(1)
        )
        tf = title_box.text_frame
        tf.text = title
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.font.name = self.FONT_NAME
        p.font.size = Pt(32)
        p.font.bold = True
        p.font.color.rgb = self.WHITE
        
        # Accent line
        line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(5.5), Inches(2.4),
            Inches(1), Inches(0.03)
        )
        line.fill.solid()
        line.fill.fore_color.rgb = self.WHITE
        line.line.fill.background()
        
        # Body text
        text_box = slide.shapes.add_textbox(
            Inches(5.5), Inches(3),
            Inches(4), Inches(4)
        )
        ttf = text_box.text_frame
        ttf.word_wrap = True
        
        for i, para_text in enumerate(paragraphs):
            if i == 0:
                para = ttf.paragraphs[0]
            else:
                para = ttf.add_paragraph()
            
            para.text = para_text
            para.font.name = self.FONT_NAME
            para.font.size = Pt(14)
            para.font.color.rgb = self.WHITE
            para.space_after = Pt(12)
        
        return slide
    
    def add_market_opportunities(self, title, items):
        """
        Create market opportunities slide
        items = [{"icon": "📊", "title": "Title", "text": "Description"}, ...]
        """
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
        # Blue accent (top right)
        blue_accent = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(8.5), Inches(0),
            Inches(1.5), Inches(1.2)
        )
        blue_accent.fill.solid()
        blue_accent.fill.fore_color.rgb = self.LIGHT_BLUE
        blue_accent.line.fill.background()
        
        # Title
        title_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(1),
            Inches(3.5), Inches(1.5)
        )
        tf = title_box.text_frame
        tf.text = title
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.font.name = self.FONT_NAME
        p.font.size = Pt(44)
        p.font.bold = True
        p.font.color.rgb = self.DARK_GRAY
        
        # Add items
        start_y = Inches(3)
        spacing = Inches(2.3)
        
        for i, item in enumerate(items[:2]):
            y_pos = start_y + (i * spacing)
            
            # Red icon background
            icon_bg = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(0.5), y_pos,
                Inches(0.7), Inches(0.7)
            )
            icon_bg.fill.solid()
            icon_bg.fill.fore_color.rgb = self.RED
            icon_bg.line.fill.background()
            
            # Icon
            icon_box = slide.shapes.add_textbox(
                Inches(0.5), y_pos,
                Inches(0.7), Inches(0.7)
            )
            itf = icon_box.text_frame
            itf.text = item.get("icon", "✓")
            itf.vertical_anchor = MSO_ANCHOR.MIDDLE
            
            ip = itf.paragraphs[0]
            ip.alignment = PP_ALIGN.CENTER
            ip.font.size = Pt(28)
            
            # Title
            title_box = slide.shapes.add_textbox(
                Inches(1.4), y_pos,
                Inches(2.8), Inches(0.5)
            )
            ttf = title_box.text_frame
            ttf.text = item["title"]
            
            tp = ttf.paragraphs[0]
            tp.font.name = self.FONT_NAME
            tp.font.size = Pt(20)
            tp.font.bold = True
            tp.font.color.rgb = self.DARK_GRAY
            
            # Description
            desc_box = slide.shapes.add_textbox(
                Inches(1.4), y_pos + Inches(0.6),
                Inches(2.8), Inches(1.5)
            )
            dtf = desc_box.text_frame
            dtf.text = item["text"]
            dtf.word_wrap = True
            
            dp = dtf.paragraphs[0]
            dp.font.name = self.FONT_NAME
            dp.font.size = Pt(12)
            dp.font.color.rgb = self.DARK_GRAY
        
        return slide
    
    def save(self, filename):
        """Save presentation to file"""
        self.prs.save(filename)
        print(f"✅ Presentation saved: {filename}")