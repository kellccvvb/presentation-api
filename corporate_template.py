from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

class CorporatePresentation:
    """
    Corporate presentation template with 22 professional slide types.
    Colors: Deep Teal (#2C5F7C), Bright Red (#E31E24), Light Blue (#7BA7BC)
    Style: Modern, clean, professional layouts for sales and business presentations
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
        self.DARK_TEAL = RGBColor(30, 70, 95)
        
        self.FONT_NAME = "Calibri"
    
    # ========== EXISTING TEMPLATES (12) ==========
    
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
    
    def add_timeline(self, title, milestones):
        """
        Create timeline slide
        milestones = [{"date": "Q1 2026", "event": "Product Launch"}, ...]
        """
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
        # Title
        title_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(0.8),
            Inches(9), Inches(0.8)
        )
        tf = title_box.text_frame
        tf.text = title
        
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
        
        # Timeline line
        timeline_line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(1), Inches(3.5),
            Inches(8), Inches(0.1)
        )
        timeline_line.fill.solid()
        timeline_line.fill.fore_color.rgb = self.TEAL
        timeline_line.line.fill.background()
        
        # Add milestones
        num_milestones = len(milestones[:5])
        spacing = 8.0 / (num_milestones + 1)
        
        for i, milestone in enumerate(milestones[:5]):
            x_pos = Inches(1 + spacing * (i + 1))
            
            # Milestone circle
            circle = slide.shapes.add_shape(
                MSO_SHAPE.OVAL,
                x_pos - Inches(0.25), Inches(3.25),
                Inches(0.5), Inches(0.5)
            )
            circle.fill.solid()
            circle.fill.fore_color.rgb = self.RED if i % 2 == 0 else self.TEAL
            circle.line.fill.background()
            
            # Date (above line)
            date_box = slide.shapes.add_textbox(
                x_pos - Inches(0.75), Inches(2.3),
                Inches(1.5), Inches(0.4)
            )
            dtf = date_box.text_frame
            dtf.text = milestone["date"]
            
            dp = dtf.paragraphs[0]
            dp.alignment = PP_ALIGN.CENTER
            dp.font.name = self.FONT_NAME
            dp.font.size = Pt(14)
            dp.font.bold = True
            dp.font.color.rgb = self.DARK_GRAY
            
            # Event (below line)
            event_box = slide.shapes.add_textbox(
                x_pos - Inches(0.75), Inches(4),
                Inches(1.5), Inches(1.5)
            )
            etf = event_box.text_frame
            etf.text = milestone["event"]
            etf.word_wrap = True
            
            ep = etf.paragraphs[0]
            ep.alignment = PP_ALIGN.CENTER
            ep.font.name = self.FONT_NAME
            ep.font.size = Pt(12)
            ep.font.color.rgb = self.DARK_GRAY
        
        return slide
    
    def add_comparison(self, title, left_side, right_side):
        """
        Create comparison slide
        left_side = {"title": "Option A", "items": ["Point 1", "Point 2"]}
        right_side = {"title": "Option B", "items": ["Point 1", "Point 2"]}
        """
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
        # Title
        title_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(0.8),
            Inches(9), Inches(0.8)
        )
        tf = title_box.text_frame
        tf.text = title
        
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
        
        # Divider line (center)
        divider = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(5), Inches(2.2),
            Inches(0.05), Inches(5)
        )
        divider.fill.solid()
        divider.fill.fore_color.rgb = self.LIGHT_GRAY
        divider.line.fill.background()
        
        # LEFT SIDE
        left_header = slide.shapes.add_textbox(
            Inches(0.5), Inches(2.2),
            Inches(4), Inches(0.6)
        )
        lhf = left_header.text_frame
        lhf.text = left_side["title"]
        
        lhp = lhf.paragraphs[0]
        lhp.font.name = self.FONT_NAME
        lhp.font.size = Pt(24)
        lhp.font.bold = True
        lhp.font.color.rgb = self.TEAL
        
        # Left items
        for i, item in enumerate(left_side["items"][:5]):
            item_box = slide.shapes.add_textbox(
                Inches(0.5), Inches(3 + i * 0.8),
                Inches(4), Inches(0.7)
            )
            itf = item_box.text_frame
            itf.text = f"• {item}"
            itf.word_wrap = True
            
            itp = itf.paragraphs[0]
            itp.font.name = self.FONT_NAME
            itp.font.size = Pt(14)
            itp.font.color.rgb = self.DARK_GRAY
        
        # RIGHT SIDE
        right_header = slide.shapes.add_textbox(
            Inches(5.5), Inches(2.2),
            Inches(4), Inches(0.6)
        )
        rhf = right_header.text_frame
        rhf.text = right_side["title"]
        
        rhp = rhf.paragraphs[0]
        rhp.font.name = self.FONT_NAME
        rhp.font.size = Pt(24)
        rhp.font.bold = True
        rhp.font.color.rgb = self.RED
        
        # Right items
        for i, item in enumerate(right_side["items"][:5]):
            item_box = slide.shapes.add_textbox(
                Inches(5.5), Inches(3 + i * 0.8),
                Inches(4), Inches(0.7)
            )
            itf = item_box.text_frame
            itf.text = f"• {item}"
            itf.word_wrap = True
            
            itp = itf.paragraphs[0]
            itp.font.name = self.FONT_NAME
            itp.font.size = Pt(14)
            itp.font.color.rgb = self.DARK_GRAY
        
        return slide
    
    def add_process_steps(self, title, steps):
        """
        Create process/steps slide
        steps = [{"number": "1", "title": "First Step", "text": "Description"}, ...]
        """
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
        # Title
        title_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(0.8),
            Inches(9), Inches(0.8)
        )
        tf = title_box.text_frame
        tf.text = title
        
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
        
        # Steps
        num_steps = len(steps[:4])
        start_y = Inches(2.5)
        spacing = Inches(1.3)
        
        for i, step in enumerate(steps[:4]):
            y_pos = start_y + (i * spacing)
            
            # Number circle
            circle = slide.shapes.add_shape(
                MSO_SHAPE.OVAL,
                Inches(0.5), y_pos,
                Inches(0.8), Inches(0.8)
            )
            circle.fill.solid()
            circle.fill.fore_color.rgb = self.TEAL
            circle.line.fill.background()
            
            # Number text
            num_box = slide.shapes.add_textbox(
                Inches(0.5), y_pos,
                Inches(0.8), Inches(0.8)
            )
            ntf = num_box.text_frame
            ntf.text = str(step.get("number", i + 1))
            ntf.vertical_anchor = MSO_ANCHOR.MIDDLE
            
            np = ntf.paragraphs[0]
            np.alignment = PP_ALIGN.CENTER
            np.font.name = self.FONT_NAME
            np.font.size = Pt(28)
            np.font.bold = True
            np.font.color.rgb = self.WHITE
            
            # Step title
            step_title = slide.shapes.add_textbox(
                Inches(1.5), y_pos,
                Inches(8), Inches(0.4)
            )
            stf = step_title.text_frame
            stf.text = step["title"]
            
            stp = stf.paragraphs[0]
            stp.font.name = self.FONT_NAME
            stp.font.size = Pt(18)
            stp.font.bold = True
            stp.font.color.rgb = self.DARK_GRAY
            
            # Step description
            step_desc = slide.shapes.add_textbox(
                Inches(1.5), y_pos + Inches(0.45),
                Inches(8), Inches(0.7)
            )
            sdf = step_desc.text_frame
            sdf.text = step["text"]
            sdf.word_wrap = True
            
            sdp = sdf.paragraphs[0]
            sdp.font.name = self.FONT_NAME
            sdp.font.size = Pt(12)
            sdp.font.color.rgb = self.DARK_GRAY
            
            # Arrow
            if i < num_steps - 1:
                arrow = slide.shapes.add_shape(
                    MSO_SHAPE.DOWN_ARROW,
                    Inches(0.65), y_pos + Inches(0.9),
                    Inches(0.5), Inches(0.3)
                )
                arrow.fill.solid()
                arrow.fill.fore_color.rgb = self.RED
                arrow.line.fill.background()
        
        return slide
    
    def add_team_slide(self, title, members):
        """
        Create team/people slide
        members = [{"name": "John Doe", "role": "CEO", "icon": "👤"}, ...]
        """
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
        # Title
        title_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(0.8),
            Inches(9), Inches(0.8)
        )
        tf = title_box.text_frame
        tf.text = title
        
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
        
        # Grid layout
        positions = [
            (Inches(1), Inches(2.5)),
            (Inches(4), Inches(2.5)),
            (Inches(7), Inches(2.5)),
            (Inches(1), Inches(5)),
            (Inches(4), Inches(5)),
            (Inches(7), Inches(5))
        ]
        
        for i, member in enumerate(members[:6]):
            if i >= len(positions):
                break
                
            x_pos, y_pos = positions[i]
            
            # Icon/avatar circle
            circle = slide.shapes.add_shape(
                MSO_SHAPE.OVAL,
                x_pos + Inches(0.5), y_pos,
                Inches(1), Inches(1)
            )
            circle.fill.solid()
            circle.fill.fore_color.rgb = self.TEAL
            circle.line.fill.background()
            
            # Icon
            icon_box = slide.shapes.add_textbox(
                x_pos + Inches(0.5), y_pos,
                Inches(1), Inches(1)
            )
            itf = icon_box.text_frame
            itf.text = member.get("icon", "👤")
            itf.vertical_anchor = MSO_ANCHOR.MIDDLE
            
            ip = itf.paragraphs[0]
            ip.alignment = PP_ALIGN.CENTER
            ip.font.size = Pt(32)
            
            # Name
            name_box = slide.shapes.add_textbox(
                x_pos, y_pos + Inches(1.1),
                Inches(2), Inches(0.3)
            )
            nf = name_box.text_frame
            nf.text = member["name"]
            
            np = nf.paragraphs[0]
            np.alignment = PP_ALIGN.CENTER
            np.font.name = self.FONT_NAME
            np.font.size = Pt(14)
            np.font.bold = True
            np.font.color.rgb = self.DARK_GRAY
            
            # Role
            role_box = slide.shapes.add_textbox(
                x_pos, y_pos + Inches(1.4),
                Inches(2), Inches(0.3)
            )
            rf = role_box.text_frame
            rf.text = member["role"]
            
            rp = rf.paragraphs[0]
            rp.alignment = PP_ALIGN.CENTER
            rp.font.name = self.FONT_NAME
            rp.font.size = Pt(11)
            rp.font.color.rgb = self.TEAL
        
        return slide
    
    def add_quote_slide(self, quote, author, role=""):
        """Create quote/testimonial slide"""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
        # Background accent
        accent = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(0), Inches(0),
            Inches(2), Inches(7.5)
        )
        accent.fill.solid()
        accent.fill.fore_color.rgb = self.TEAL
        accent.line.fill.background()
        
        # Quote marks
        quote_mark1 = slide.shapes.add_textbox(
            Inches(2.5), Inches(2),
            Inches(1), Inches(1)
        )
        qm1f = quote_mark1.text_frame
        qm1f.text = '"'
        
        qm1p = qm1f.paragraphs[0]
        qm1p.font.name = self.FONT_NAME
        qm1p.font.size = Pt(120)
        qm1p.font.color.rgb = self.RED
        
        # Quote text
        quote_box = slide.shapes.add_textbox(
            Inches(2.5), Inches(3),
            Inches(6.5), Inches(3)
        )
        qf = quote_box.text_frame
        qf.text = quote
        qf.word_wrap = True
        
        qp = qf.paragraphs[0]
        qp.font.name = self.FONT_NAME
        qp.font.size = Pt(24)
        qp.font.italic = True
        qp.font.color.rgb = self.DARK_GRAY
        
        # Author
        author_box = slide.shapes.add_textbox(
            Inches(2.5), Inches(6),
            Inches(6.5), Inches(0.4)
        )
        af = author_box.text_frame
        af.text = f"— {author}"
        
        ap = af.paragraphs[0]
        ap.font.name = self.FONT_NAME
        ap.font.size = Pt(18)
        ap.font.bold = True
        ap.font.color.rgb = self.TEAL
        
        # Role
        if role:
            role_box = slide.shapes.add_textbox(
                Inches(2.5), Inches(6.4),
                Inches(6.5), Inches(0.3)
            )
            rf = role_box.text_frame
            rf.text = role
            
            rp = rf.paragraphs[0]
            rp.font.name = self.FONT_NAME
            rp.font.size = Pt(14)
            rp.font.color.rgb = self.DARK_GRAY
        
        return slide
    
    def add_stats_slide(self, title, stats):
        """
        Create data/stats slide with big numbers
        stats = [{"number": "25%", "label": "Revenue Growth"}, ...]
        """
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
        # Title
        title_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(0.8),
            Inches(9), Inches(0.8)
        )
        tf = title_box.text_frame
        tf.text = title
        
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
        
        # Grid layout
        positions = [
            (Inches(1.5), Inches(3)),
            (Inches(5.5), Inches(3)),
            (Inches(1.5), Inches(5.5)),
            (Inches(5.5), Inches(5.5))
        ]
        
        colors = [self.TEAL, self.RED, self.LIGHT_BLUE, self.TEAL]
        
        for i, stat in enumerate(stats[:4]):
            if i >= len(positions):
                break
                
            x_pos, y_pos = positions[i]
            
            # Number
            num_box = slide.shapes.add_textbox(
                x_pos, y_pos,
                Inches(3), Inches(1)
            )
            nf = num_box.text_frame
            nf.text = stat["number"]
            
            np = nf.paragraphs[0]
            np.alignment = PP_ALIGN.CENTER
            np.font.name = self.FONT_NAME
            np.font.size = Pt(54)
            np.font.bold = True
            np.font.color.rgb = colors[i]
            
            # Label
            label_box = slide.shapes.add_textbox(
                x_pos, y_pos + Inches(1),
                Inches(3), Inches(0.6)
            )
            lf = label_box.text_frame
            lf.text = stat["label"]
            lf.word_wrap = True
            
            lp = lf.paragraphs[0]
            lp.alignment = PP_ALIGN.CENTER
            lp.font.name = self.FONT_NAME
            lp.font.size = Pt(14)
            lp.font.color.rgb = self.DARK_GRAY
        
        return slide
    
    def add_chart_slide(self, title, chart_data, chart_type="bar"):
        """
        Create chart slide
        chart_data = {
            "categories": ["Q1", "Q2", "Q3", "Q4"],
            "series": [
                {"name": "Revenue", "values": [10, 15, 20, 25]},
                {"name": "Profit", "values": [5, 8, 12, 15]}
            ]
        }
        """
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
        # Title
        title_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(0.8),
            Inches(9), Inches(0.8)
        )
        tf = title_box.text_frame
        tf.text = title
        
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
        
        # Chart data
        chart_data_obj = CategoryChartData()
        chart_data_obj.categories = chart_data["categories"]
        
        for series in chart_data["series"]:
            chart_data_obj.add_series(series["name"], series["values"])
        
        # Add chart
        x, y, cx, cy = Inches(1), Inches(2.5), Inches(8), Inches(4.5)
        
        if chart_type == "line":
            chart = slide.shapes.add_chart(
                XL_CHART_TYPE.LINE,
                x, y, cx, cy,
                chart_data_obj
            ).chart
        else:
            chart = slide.shapes.add_chart(
                XL_CHART_TYPE.COLUMN_CLUSTERED,
                x, y, cx, cy,
                chart_data_obj
            ).chart
        
        return slide
    
    # ========== NEW TEMPLATES (10) ==========
    
    def add_chart_with_icons(self, title, chart_data, advantages):
        """
        NEW: Chart + icon list combo (like Competitive Advantages)
        chart_data = {"categories": ["Item 1", "Item 2"], "values": [10, 15]}
        advantages = [{"icon": "🌐", "title": "Title", "text": "Description"}, ...]
        """
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
        # Title (left side)
        title_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(0.8),
            Inches(4), Inches(1.2)
        )
        tf = title_box.text_frame
        tf.text = title
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.font.name = self.FONT_NAME
        p.font.size = Pt(48)
        p.font.bold = True
        p.font.color.rgb = self.DARK_GRAY
        
        # Bar chart (left side)
        chart_data_obj = CategoryChartData()
        chart_data_obj.categories = chart_data["categories"]
        chart_data_obj.add_series('Data', chart_data["values"])
        
        x, y, cx, cy = Inches(0.5), Inches(2.5), Inches(4), Inches(4.5)
        chart = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            x, y, cx, cy,
            chart_data_obj
        ).chart
        
        # Icon list (right side)
        start_y = Inches(1.5)
        spacing = Inches(1.5)
        
        for i, adv in enumerate(advantages[:4]):
            y_pos = start_y + (i * spacing)
            
            # Red icon box
            icon_bg = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(5), y_pos,
                Inches(0.7), Inches(0.7)
            )
            icon_bg.fill.solid()
            icon_bg.fill.fore_color.rgb = self.RED
            icon_bg.line.fill.background()
            
            # Icon
            icon_box = slide.shapes.add_textbox(
                Inches(5), y_pos,
                Inches(0.7), Inches(0.7)
            )
            itf = icon_box.text_frame
            itf.text = adv.get("icon", "✓")
            itf.vertical_anchor = MSO_ANCHOR.MIDDLE
            
            ip = itf.paragraphs[0]
            ip.alignment = PP_ALIGN.CENTER
            ip.font.size = Pt(28)
            
            # Title
            title_box = slide.shapes.add_textbox(
                Inches(5.9), y_pos,
                Inches(3.6), Inches(0.35)
            )
            ttf = title_box.text_frame
            ttf.text = adv["title"]
            
            tp = ttf.paragraphs[0]
            tp.font.name = self.FONT_NAME
            tp.font.size = Pt(18)
            tp.font.bold = True
            tp.font.color.rgb = self.DARK_GRAY
            
            # Description
            desc_box = slide.shapes.add_textbox(
                Inches(5.9), y_pos + Inches(0.4),
                Inches(3.6), Inches(0.9)
            )
            dtf = desc_box.text_frame
            dtf.text = adv["text"]
            dtf.word_wrap = True
            
            dp = dtf.paragraphs[0]
            dp.font.name = self.FONT_NAME
            dp.font.size = Pt(11)
            dp.font.color.rgb = self.DARK_GRAY
        
        return slide
    
    def add_image_with_sidebar(self, title, subtitle, checklist):
        """
        NEW: Image + sidebar + checklist (like Best Platform slide)
        checklist = ["Item 1", "Item 2"]
        """
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
        # Blue sidebar (left)
        sidebar = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(0), Inches(0),
            Inches(0.8), Inches(7.5)
        )
        sidebar.fill.solid()
        sidebar.fill.fore_color.rgb = self.TEAL
        sidebar.line.fill.background()
        
        # Sidebar icon/label (rotated text effect with icon)
        icon_box = slide.shapes.add_textbox(
            Inches(0.1), Inches(3),
            Inches(0.6), Inches(2)
        )
        itf = icon_box.text_frame
        itf.text = "📊"
        itf.vertical_anchor = MSO_ANCHOR.MIDDLE
        
        ip = itf.paragraphs[0]
        ip.alignment = PP_ALIGN.CENTER
        ip.font.size = Pt(32)
        
        # Title (right side)
        title_box = slide.shapes.add_textbox(
            Inches(5), Inches(1.5),
            Inches(4.5), Inches(1.5)
        )
        tf = title_box.text_frame
        tf.text = title
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.font.name = self.FONT_NAME
        p.font.size = Pt(44)
        p.font.bold = True
        p.font.color.rgb = self.DARK_GRAY
        
        # Red accent line
        line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(5), Inches(3),
            Inches(1.5), Inches(0.08)
        )
        line.fill.solid()
        line.fill.fore_color.rgb = self.RED
        line.line.fill.background()
        
        # Subtitle
        subtitle_box = slide.shapes.add_textbox(
            Inches(5), Inches(3.3),
            Inches(4.5), Inches(1)
        )
        stf = subtitle_box.text_frame
        stf.text = subtitle
        stf.word_wrap = True
        
        sp = stf.paragraphs[0]
        sp.font.name = self.FONT_NAME
        sp.font.size = Pt(14)
        sp.font.color.rgb = self.DARK_GRAY
        
        # Checklist items
        start_y = Inches(4.5)
        for i, item in enumerate(checklist[:3]):
            check_box = slide.shapes.add_textbox(
                Inches(5), start_y + (i * Inches(0.6)),
                Inches(4.5), Inches(0.5)
            )
            ctf = check_box.text_frame
            ctf.text = f"✓ {item}"
            
            cp = ctf.paragraphs[0]
            cp.font.name = self.FONT_NAME
            cp.font.size = Pt(14)
            cp.font.color.rgb = self.DARK_GRAY
        
        return slide
    
    def add_pricing_table(self, title, plans):
        """
        NEW: Pricing table with 3 tiers
        plans = [
            {"name": "Starter", "price": "$29/mo", "features": ["Feature 1", "Feature 2"], "highlighted": False},
            {"name": "Professional", "price": "$99/mo", "features": ["Feature 1", "Feature 2"], "highlighted": True},
            {"name": "Enterprise", "price": "$299/mo", "features": ["Feature 1", "Feature 2"], "highlighted": False}
        ]
        """
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
        # Title
        title_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(0.8),
            Inches(9), Inches(0.8)
        )
        tf = title_box.text_frame
        tf.text = title
        
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
        
        # Three pricing columns
        x_positions = [Inches(0.8), Inches(3.7), Inches(6.6)]
        
        for i, plan in enumerate(plans[:3]):
            x_pos = x_positions[i]
            
            # Plan box
            box_color = self.RED if plan.get("highlighted") else self.LIGHT_GRAY
            plan_box = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                x_pos, Inches(2.3),
                Inches(2.5), Inches(4.7)
            )
            plan_box.fill.solid()
            plan_box.fill.fore_color.rgb = box_color
            plan_box.line.color.rgb = self.DARK_GRAY
            plan_box.line.width = Pt(1)
            
            # Plan name
            name_box = slide.shapes.add_textbox(
                x_pos + Inches(0.2), Inches(2.6),
                Inches(2.1), Inches(0.5)
            )
            nf = name_box.text_frame
            nf.text = plan["name"]
            
            np = nf.paragraphs[0]
            np.alignment = PP_ALIGN.CENTER
            np.font.name = self.FONT_NAME
            np.font.size = Pt(24)
            np.font.bold = True
            np.font.color.rgb = self.WHITE if plan.get("highlighted") else self.DARK_GRAY
            
            # Price
            price_box = slide.shapes.add_textbox(
                x_pos + Inches(0.2), Inches(3.2),
                Inches(2.1), Inches(0.5)
            )
            pf = price_box.text_frame
            pf.text = plan["price"]
            
            pp = pf.paragraphs[0]
            pp.alignment = PP_ALIGN.CENTER
            pp.font.name = self.FONT_NAME
            pp.font.size = Pt(32)
            pp.font.bold = True
            pp.font.color.rgb = self.WHITE if plan.get("highlighted") else self.TEAL
            
            # Features
            features_box = slide.shapes.add_textbox(
                x_pos + Inches(0.2), Inches(4),
                Inches(2.1), Inches(2.5)
            )
            ff = features_box.text_frame
            ff.word_wrap = True
            
            for j, feature in enumerate(plan["features"][:5]):
                if j == 0:
                    para = ff.paragraphs[0]
                else:
                    para = ff.add_paragraph()
                
                para.text = f"✓ {feature}"
                para.font.name = self.FONT_NAME
                para.font.size = Pt(11)
                para.font.color.rgb = self.WHITE if plan.get("highlighted") else self.DARK_GRAY
                para.space_after = Pt(6)
        
        return slide
    
    def add_feature_comparison_table(self, title, features, competitors):
        """
        NEW: Feature comparison table (You vs Competitor A vs Competitor B)
        features = ["Feature 1", "Feature 2", "Feature 3"]
        competitors = [
            {"name": "Us", "has_features": [True, True, True]},
            {"name": "Competitor A", "has_features": [True, False, True]},
            {"name": "Competitor B", "has_features": [False, True, False]}
        ]
        """
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
        # Title
        title_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(0.8),
            Inches(9), Inches(0.8)
        )
        tf = title_box.text_frame
        tf.text = title
        
        p = tf.paragraphs[0]
        p.font.name = self.FONT_NAME
        p.font.size = Pt(40)
        p.font.bold = True
        p.font.color.rgb = self.DARK_GRAY
        
        # Table headers (competitor names)
        header_y = Inches(2)
        col_width = Inches(2.5)
        
        for i, comp in enumerate(competitors[:3]):
            header_box = slide.shapes.add_textbox(
                Inches(2.5 + i * col_width), header_y,
                col_width, Inches(0.5)
            )
            hf = header_box.text_frame
            hf.text = comp["name"]
            
            hp = hf.paragraphs[0]
            hp.alignment = PP_ALIGN.CENTER
            hp.font.name = self.FONT_NAME
            hp.font.size = Pt(18)
            hp.font.bold = True
            hp.font.color.rgb = self.TEAL if i == 0 else self.DARK_GRAY
        
        # Feature rows
        row_height = Inches(0.7)
        start_y = Inches(2.8)
        
        for i, feature in enumerate(features[:6]):
            y_pos = start_y + (i * row_height)
            
            # Feature name
            feature_box = slide.shapes.add_textbox(
                Inches(0.5), y_pos,
                Inches(1.8), row_height
            )
            ff = feature_box.text_frame
            ff.text = feature
            ff.word_wrap = True
            ff.vertical_anchor = MSO_ANCHOR.MIDDLE
            
            fp = ff.paragraphs[0]
            fp.font.name = self.FONT_NAME
            fp.font.size = Pt(12)
            fp.font.color.rgb = self.DARK_GRAY
            
            # Checkmarks for each competitor
            for j, comp in enumerate(competitors[:3]):
                has_feature = comp["has_features"][i] if i < len(comp["has_features"]) else False
                
                check_box = slide.shapes.add_textbox(
                    Inches(2.5 + j * col_width), y_pos,
                    col_width, row_height
                )
                cf = check_box.text_frame
                cf.text = "✓" if has_feature else "✗"
                cf.vertical_anchor = MSO_ANCHOR.MIDDLE
                
                cp = cf.paragraphs[0]
                cp.alignment = PP_ALIGN.CENTER
                cp.font.name = self.FONT_NAME
                cp.font.size = Pt(24)
                cp.font.color.rgb = self.TEAL if has_feature else self.LIGHT_GRAY
        
        return slide
    
    def add_roi_calculator(self, title, current_cost, solution_cost, savings, payback_months):
        """
        NEW: ROI/Value calculator slide
        Shows cost comparison and ROI metrics
        """
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
        # Title
        title_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(0.8),
            Inches(9), Inches(0.8)
        )
        tf = title_box.text_frame
        tf.text = title
        
        p = tf.paragraphs[0]
        p.font.name = self.FONT_NAME
        p.font.size = Pt(40)
        p.font.bold = True
        p.font.color.rgb = self.DARK_GRAY
        
        # Current cost box
        current_box = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(1), Inches(2.5),
            Inches(3.5), Inches(1.5)
        )
        current_box.fill.solid()
        current_box.fill.fore_color.rgb = self.LIGHT_GRAY
        current_box.line.fill.background()
        
        current_label = slide.shapes.add_textbox(
            Inches(1.2), Inches(2.7),
            Inches(3.1), Inches(0.4)
        )
        clf = current_label.text_frame
        clf.text = "Current Annual Cost"
        
        clp = clf.paragraphs[0]
        clp.alignment = PP_ALIGN.CENTER
        clp.font.name = self.FONT_NAME
        clp.font.size = Pt(14)
        clp.font.color.rgb = self.DARK_GRAY
        
        current_value = slide.shapes.add_textbox(
            Inches(1.2), Inches(3.2),
            Inches(3.1), Inches(0.6)
        )
        cvf = current_value.text_frame
        cvf.text = current_cost
        
        cvp = cvf.paragraphs[0]
        cvp.alignment = PP_ALIGN.CENTER
        cvp.font.name = self.FONT_NAME
        cvp.font.size = Pt(36)
        cvp.font.bold = True
        cvp.font.color.rgb = self.RED
        
        # Solution cost box
        solution_box = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(5.5), Inches(2.5),
            Inches(3.5), Inches(1.5)
        )
        solution_box.fill.solid()
        solution_box.fill.fore_color.rgb = self.TEAL
        solution_box.line.fill.background()
        
        solution_label = slide.shapes.add_textbox(
            Inches(5.7), Inches(2.7),
            Inches(3.1), Inches(0.4)
        )
        slf = solution_label.text_frame
        slf.text = "With Our Solution"
        
        slp = slf.paragraphs[0]
        slp.alignment = PP_ALIGN.CENTER
        slp.font.name = self.FONT_NAME
        slp.font.size = Pt(14)
        slp.font.color.rgb = self.WHITE
        
        solution_value = slide.shapes.add_textbox(
            Inches(5.7), Inches(3.2),
            Inches(3.1), Inches(0.6)
        )
        svf = solution_value.text_frame
        svf.text = solution_cost
        
        svp = svf.paragraphs[0]
        svp.alignment = PP_ALIGN.CENTER
        svp.font.name = self.FONT_NAME
        svp.font.size = Pt(36)
        svp.font.bold = True
        svp.font.color.rgb = self.WHITE
        
        # Savings metrics
        savings_box = slide.shapes.add_textbox(
            Inches(2), Inches(4.5),
            Inches(6), Inches(0.6)
        )
        sbf = savings_box.text_frame
        sbf.text = f"Annual Savings: {savings}"
        
        sbp = sbf.paragraphs[0]
        sbp.alignment = PP_ALIGN.CENTER
        sbp.font.name = self.FONT_NAME
        sbp.font.size = Pt(28)
        sbp.font.bold = True
        sbp.font.color.rgb = self.TEAL
        
        payback_box = slide.shapes.add_textbox(
            Inches(2), Inches(5.5),
            Inches(6), Inches(0.5)
        )
        pbf = payback_box.text_frame
        pbf.text = f"Payback Period: {payback_months} months"
        
        pbp = pbf.paragraphs[0]
        pbp.alignment = PP_ALIGN.CENTER
        pbp.font.name = self.FONT_NAME
        pbp.font.size = Pt(20)
        pbp.font.color.rgb = self.DARK_GRAY
        
        return slide
    
    def add_case_study(self, customer_name, industry, challenge, solution, results):
        """
        NEW: Case study/success story slide
        results = [{"metric": "25%", "label": "Cost Reduction"}, ...]
        """
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
        # Title
        title_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(0.8),
            Inches(5), Inches(0.6)
        )
        tf = title_box.text_frame
        tf.text = f"Case Study: {customer_name}"
        
        p = tf.paragraphs[0]
        p.font.name = self.FONT_NAME
        p.font.size = Pt(36)
        p.font.bold = True
        p.font.color.rgb = self.DARK_GRAY
        
        # Industry tag
        industry_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(1.5),
            Inches(2), Inches(0.4)
        )
        inf = industry_box.text_frame
        inf.text = industry
        
        inp = inf.paragraphs[0]
        inp.font.name = self.FONT_NAME
        inp.font.size = Pt(14)
        inp.font.color.rgb = self.TEAL
        
        # Challenge section
        challenge_label = slide.shapes.add_textbox(
            Inches(0.5), Inches(2.2),
            Inches(4.5), Inches(0.4)
        )
        clf = challenge_label.text_frame
        clf.text = "Challenge"
        
        clp = clf.paragraphs[0]
        clp.font.name = self.FONT_NAME
        clp.font.size = Pt(18)
        clp.font.bold = True
        clp.font.color.rgb = self.RED
        
        challenge_text = slide.shapes.add_textbox(
            Inches(0.5), Inches(2.7),
            Inches(4.5), Inches(1)
        )
        ctf = challenge_text.text_frame
        ctf.text = challenge
        ctf.word_wrap = True
        
        ctp = ctf.paragraphs[0]
        ctp.font.name = self.FONT_NAME
        ctp.font.size = Pt(12)
        ctp.font.color.rgb = self.DARK_GRAY
        
        # Solution section
        solution_label = slide.shapes.add_textbox(
            Inches(0.5), Inches(4),
            Inches(4.5), Inches(0.4)
        )
        slf = solution_label.text_frame
        slf.text = "Solution"
        
        slp = slf.paragraphs[0]
        slp.font.name = self.FONT_NAME
        slp.font.size = Pt(18)
        slp.font.bold = True
        slp.font.color.rgb = self.TEAL
        
        solution_text = slide.shapes.add_textbox(
            Inches(0.5), Inches(4.5),
            Inches(4.5), Inches(1)
        )
        stf = solution_text.text_frame
        stf.text = solution
        stf.word_wrap = True
        
        stp = stf.paragraphs[0]
        stp.font.name = self.FONT_NAME
        stp.font.size = Pt(12)
        stp.font.color.rgb = self.DARK_GRAY
        
        # Results boxes (right side)
        results_label = slide.shapes.add_textbox(
            Inches(5.5), Inches(2.2),
            Inches(4), Inches(0.4)
        )
        rlf = results_label.text_frame
        rlf.text = "Results"
        
        rlp = rlf.paragraphs[0]
        rlp.font.name = self.FONT_NAME
        rlp.font.size = Pt(24)
        rlp.font.bold = True
        rlp.font.color.rgb = self.DARK_GRAY
        
        y_positions = [Inches(3), Inches(4.5), Inches(6)]
        
        for i, result in enumerate(results[:3]):
            if i >= len(y_positions):
                break
                
            y_pos = y_positions[i]
            
            # Metric
            metric_box = slide.shapes.add_textbox(
                Inches(5.5), y_pos,
                Inches(4), Inches(0.6)
            )
            mf = metric_box.text_frame
            mf.text = result["metric"]
            
            mp = mf.paragraphs[0]
            mp.alignment = PP_ALIGN.CENTER
            mp.font.name = self.FONT_NAME
            mp.font.size = Pt(40)
            mp.font.bold = True
            mp.font.color.rgb = self.TEAL
            
            # Label
            label_box = slide.shapes.add_textbox(
                Inches(5.5), y_pos + Inches(0.6),
                Inches(4), Inches(0.4)
            )
            lf = label_box.text_frame
            lf.text = result["label"]
            lf.word_wrap = True
            
            lp = lf.paragraphs[0]
            lp.alignment = PP_ALIGN.CENTER
            lp.font.name = self.FONT_NAME
            lp.font.size = Pt(12)
            lp.font.color.rgb = self.DARK_GRAY
        
        return slide
    
    def add_problem_solution(self, title, problems, solutions):
        """
        NEW: Problem-solution split slide
        problems = ["Problem 1", "Problem 2", "Problem 3"]
        solutions = ["Solution 1", "Solution 2", "Solution 3"]
        """
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
        # Title
        title_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(0.8),
            Inches(9), Inches(0.8)
        )
        tf = title_box.text_frame
        tf.text = title
        
        p = tf.paragraphs[0]
        p.font.name = self.FONT_NAME
        p.font.size = Pt(40)
        p.font.bold = True
        p.font.color.rgb = self.DARK_GRAY
        
        # Left side (Problems) - Red background
        left_panel = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(0), Inches(2),
            Inches(5), Inches(5.5)
        )
        left_panel.fill.solid()
        left_panel.fill.fore_color.rgb = RGBColor(220, 50, 50)
        left_panel.line.fill.background()
        
        # Problems header
        prob_header = slide.shapes.add_textbox(
            Inches(0.5), Inches(2.3),
            Inches(4), Inches(0.5)
        )
        phf = prob_header.text_frame
        phf.text = "Challenges"
        
        php = phf.paragraphs[0]
        php.font.name = self.FONT_NAME
        php.font.size = Pt(28)
        php.font.bold = True
        php.font.color.rgb = self.WHITE
        
        # Problems list
        for i, problem in enumerate(problems[:4]):
            prob_box = slide.shapes.add_textbox(
                Inches(0.5), Inches(3.2 + i * 0.9),
                Inches(4), Inches(0.8)
            )
            pf = prob_box.text_frame
            pf.text = f"✗ {problem}"
            pf.word_wrap = True
            
            pp = pf.paragraphs[0]
            pp.font.name = self.FONT_NAME
            pp.font.size = Pt(14)
            pp.font.color.rgb = self.WHITE
        
        # Right side (Solutions) - Teal background
        right_panel = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(5), Inches(2),
            Inches(5), Inches(5.5)
        )
        right_panel.fill.solid()
        right_panel.fill.fore_color.rgb = self.TEAL
        right_panel.line.fill.background()
        
        # Solutions header
        sol_header = slide.shapes.add_textbox(
            Inches(5.5), Inches(2.3),
            Inches(4), Inches(0.5)
        )
        shf = sol_header.text_frame
        shf.text = "Our Solutions"
        
        shp = shf.paragraphs[0]
        shp.font.name = self.FONT_NAME
        shp.font.size = Pt(28)
        shp.font.bold = True
        shp.font.color.rgb = self.WHITE
        
        # Solutions list
        for i, solution in enumerate(solutions[:4]):
            sol_box = slide.shapes.add_textbox(
                Inches(5.5), Inches(3.2 + i * 0.9),
                Inches(4), Inches(0.8)
            )
            sf = sol_box.text_frame
            sf.text = f"✓ {solution}"
            sf.word_wrap = True
            
            sp = sf.paragraphs[0]
            sp.font.name = self.FONT_NAME
            sp.font.size = Pt(14)
            sp.font.color.rgb = self.WHITE
        
        return slide
    
    def add_social_proof(self, title, logos, testimonial_quote="", testimonial_author=""):
        """
        NEW: Social proof grid (customer logos + optional testimonial)
        logos = ["Company A", "Company B", "Company C", ...] (up to 9)
        """
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
        # Title
        title_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(0.8),
            Inches(9), Inches(0.8)
        )
        tf = title_box.text_frame
        tf.text = title
        
        p = tf.paragraphs[0]
        p.font.name = self.FONT_NAME
        p.font.size = Pt(40)
        p.font.bold = True
        p.font.color.rgb = self.DARK_GRAY
        
        # Logo grid (3x3)
        positions = [
            (Inches(1), Inches(2.2)),
            (Inches(3.7), Inches(2.2)),
            (Inches(6.4), Inches(2.2)),
            (Inches(1), Inches(3.5)),
            (Inches(3.7), Inches(3.5)),
            (Inches(6.4), Inches(3.5)),
            (Inches(1), Inches(4.8)),
            (Inches(3.7), Inches(4.8)),
            (Inches(6.4), Inches(4.8))
        ]
        
        for i, logo in enumerate(logos[:9]):
            if i >= len(positions):
                break
                
            x_pos, y_pos = positions[i]
            
            logo_box = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                x_pos, y_pos,
                Inches(2.3), Inches(1)
            )
            logo_box.fill.solid()
            logo_box.fill.fore_color.rgb = self.LIGHT_GRAY
            logo_box.line.color.rgb = self.DARK_GRAY
            logo_box.line.width = Pt(0.5)
            
            text_box = slide.shapes.add_textbox(
                x_pos + Inches(0.1), y_pos + Inches(0.2),
                Inches(2.1), Inches(0.6)
            )
            tbf = text_box.text_frame
            tbf.text = logo
            tbf.vertical_anchor = MSO_ANCHOR.MIDDLE
            
            tbp = tbf.paragraphs[0]
            tbp.alignment = PP_ALIGN.CENTER
            tbp.font.name = self.FONT_NAME
            tbp.font.size = Pt(16)
            tbp.font.bold = True
            tbp.font.color.rgb = self.TEAL
        
        # Optional testimonial
        if testimonial_quote:
            quote_box = slide.shapes.add_textbox(
                Inches(1), Inches(6.2),
                Inches(8), Inches(0.8)
            )
            qf = quote_box.text_frame
            qf.text = f'"{testimonial_quote}"'
            qf.word_wrap = True
            
            qp = qf.paragraphs[0]
            qp.alignment = PP_ALIGN.CENTER
            qp.font.name = self.FONT_NAME
            qp.font.size = Pt(14)
            qp.font.italic = True
            qp.font.color.rgb = self.DARK_GRAY
            
            if testimonial_author:
                author_box = slide.shapes.add_textbox(
                    Inches(1), Inches(7),
                    Inches(8), Inches(0.3)
                )
                af = author_box.text_frame
                af.text = f"— {testimonial_author}"
                
                ap = af.paragraphs[0]
                ap.alignment = PP_ALIGN.CENTER
                ap.font.name = self.FONT_NAME
                ap.font.size = Pt(12)
                ap.font.color.rgb = self.TEAL
        
        return slide
    
    def add_product_showcase(self, title, features):
        """
        NEW: Product showcase (3x3 grid of features)
        features = [{"icon": "📱", "title": "Feature", "text": "Description"}, ...]
        """
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
        # Title
        title_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(0.8),
            Inches(9), Inches(0.8)
        )
        tf = title_box.text_frame
        tf.text = title
        
        p = tf.paragraphs[0]
        p.font.name = self.FONT_NAME
        p.font.size = Pt(40)
        p.font.bold = True
        p.font.color.rgb = self.DARK_GRAY
        
        # Grid positions (3x3)
        positions = [
            (Inches(0.5), Inches(2)),
            (Inches(3.5), Inches(2)),
            (Inches(6.5), Inches(2)),
            (Inches(0.5), Inches(4)),
            (Inches(3.5), Inches(4)),
            (Inches(6.5), Inches(4)),
            (Inches(0.5), Inches(6)),
            (Inches(3.5), Inches(6)),
            (Inches(6.5), Inches(6))
        ]
        
        for i, feature in enumerate(features[:9]):
            if i >= len(positions):
                break
                
            x_pos, y_pos = positions[i]
            
            # Icon
            icon_box = slide.shapes.add_textbox(
                x_pos + Inches(0.8), y_pos,
                Inches(1), Inches(0.5)
            )
            itf = icon_box.text_frame
            itf.text = feature.get("icon", "✓")
            itf.vertical_anchor = MSO_ANCHOR.MIDDLE
            
            ip = itf.paragraphs[0]
            ip.alignment = PP_ALIGN.CENTER
            ip.font.size = Pt(36)
            
            # Title
            title_box = slide.shapes.add_textbox(
                x_pos, y_pos + Inches(0.5),
                Inches(2.6), Inches(0.35)
            )
            ttf = title_box.text_frame
            ttf.text = feature["title"]
            ttf.word_wrap = True
            
            tp = ttf.paragraphs[0]
            tp.alignment = PP_ALIGN.CENTER
            tp.font.name = self.FONT_NAME
            tp.font.size = Pt(14)
            tp.font.bold = True
            tp.font.color.rgb = self.DARK_GRAY
            
            # Description
            desc_box = slide.shapes.add_textbox(
                x_pos, y_pos + Inches(0.9),
                Inches(2.6), Inches(0.9)
            )
            dtf = desc_box.text_frame
            dtf.text = feature["text"]
            dtf.word_wrap = True
            
            dp = dtf.paragraphs[0]
            dp.alignment = PP_ALIGN.CENTER
            dp.font.name = self.FONT_NAME
            dp.font.size = Pt(10)
            dp.font.color.rgb = self.DARK_GRAY
        
        return slide
    
    def add_pie_chart(self, title, data):
        """
        NEW: Pie/donut chart
        data = [{"label": "Category A", "value": 40}, {"label": "Category B", "value": 60}]
        """
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
        # Title
        title_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(0.8),
            Inches(9), Inches(0.8)
        )
        tf = title_box.text_frame
        tf.text = title
        
        p = tf.paragraphs[0]
        p.font.name = self.FONT_NAME
        p.font.size = Pt(40)
        p.font.bold = True
        p.font.color.rgb = self.DARK_GRAY
        
        # Chart data
        chart_data_obj = CategoryChartData()
        chart_data_obj.categories = [item["label"] for item in data]
        chart_data_obj.add_series('Data', [item["value"] for item in data])
        
        # Add pie chart
        x, y, cx, cy = Inches(2), Inches(2.5), Inches(6), Inches(4.5)
        chart = slide.shapes.add_chart(
            XL_CHART_TYPE.PIE,
            x, y, cx, cy,
            chart_data_obj
        ).chart
        
        return slide
    
    def save(self, filename):
        """Save presentation to file"""
        self.prs.save(filename)
        print(f"✅ Presentation saved: {filename}")

