from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

class CorporatePresentation:
    """
    Corporate presentation template with comprehensive slide types.
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
    
    # ========== NEW SLIDE TYPE 1: TIMELINE ==========
    
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
        num_milestones = len(milestones[:5])  # Max 5 milestones
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
    
    # ========== NEW SLIDE TYPE 2: COMPARISON ==========
    
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
    
    # ========== NEW SLIDE TYPE 3: PROCESS/STEPS ==========
    
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
        
        # Steps (max 4)
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
            
            # Arrow (if not last step)
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
    
    # ========== NEW SLIDE TYPE 4: TEAM/PEOPLE ==========
    
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
        
        # Grid layout (2x3 = 6 members max)
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
    
    # ========== NEW SLIDE TYPE 5: QUOTE/TESTIMONIAL ==========
    
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
        
        # Quote marks (top)
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
        
        # Role (if provided)
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
    
    # ========== NEW SLIDE TYPE 6: DATA/STATS ==========
    
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
        
        # Grid layout for stats (max 4)
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
    
    # ========== NEW SLIDE TYPE 7: CHART ==========
    
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
        chart_type = "bar" or "line"
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
        else:  # bar
            chart = slide.shapes.add_chart(
                XL_CHART_TYPE.COLUMN_CLUSTERED,
                x, y, cx, cy,
                chart_data_obj
            ).chart
        
        return slide
    
    def save(self, filename):
        """Save presentation to file"""
        self.prs.save(filename)
        print(f"✅ Presentation saved: {filename}")
