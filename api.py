from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from corporate_template import CorporatePresentation
import json
import os
from datetime import datetime
import io

app = Flask(__name__)
CORS(app)  # Allow requests from Base44

@app.route('/create-presentation', methods=['POST'])
def create_presentation():
    """
    Expects JSON like:
    {
        "title": "Q4 Review",
        "subtitle": "Business Update 2026",
        "sections": ["Revenue", "Growth", "Plans"],
        "slides": [
            {
                "type": "content_with_icons",
                "title": "Achievements",
                "items": [
                    {"icon": "📈", "text": "Revenue up 25%"},
                    {"icon": "🎯", "text": "All targets hit"}
                ],
                "slide_number": 3
            },
            {
                "type": "split",
                "title": "Business Plan",
                "paragraphs": ["paragraph 1", "paragraph 2"],
                "slide_number": 4
            }
        ]
    }
    """
    try:
        data = request.json
        
        # Create presentation
        deck = CorporatePresentation()
        
        slide_counter = 1
        
        # Add title slide
        deck.add_title_slide(
            title=data.get('title', 'Presentation'),
            subtitle=data.get('subtitle', ''),
            slide_number=slide_counter
        )
        slide_counter += 1
        
        # Add table of contents if sections provided
        if 'sections' in data and data['sections']:
            deck.add_table_of_contents(
                sections=data['sections'],
                slide_number=slide_counter
            )
            slide_counter += 1
        
        # Add content slides
        for slide_data in data.get('slides', []):
            slide_type = slide_data.get('type')
            slide_number = slide_data.get('slide_number', slide_counter)
            
            if slide_type == 'content_with_icons':
                deck.add_content_with_icons_slide(
                    title=slide_data.get('title', ''),
                    items=slide_data.get('items', []),
                    slide_number=slide_number
                )
            
            elif slide_type == 'split':
                deck.add_split_slide(
                    title=slide_data.get('title', ''),
                    paragraphs=slide_data.get('paragraphs', []),
                    slide_number=slide_number
                )
            
            elif slide_type == 'market_opportunities':
                deck.add_market_opportunities_slide(
                    title=slide_data.get('title', ''),
                    items=slide_data.get('items', []),
                    slide_number=slide_number
                )
            
            elif slide_type == 'timeline':
                deck.add_timeline_slide(
                    title=slide_data.get('title', ''),
                    image_url=slide_data.get('image_url', ''),
                    milestones=slide_data.get('milestones', []),
                    slide_number=slide_number
                )
            
            elif slide_type == 'comparison':
                deck.add_comparison_slide(
                    title=slide_data.get('title', ''),
                    left_side=slide_data.get('left_side', {}),
                    right_side=slide_data.get('right_side', {}),
                    middle_side=slide_data.get('middle_side'),
                    slide_number=slide_number
                )
            
            elif slide_type == 'process_steps':
                deck.add_process_steps_slide(
                    title=slide_data.get('title', ''),
                    steps=slide_data.get('steps', []),
                    slide_number=slide_number
                )
            
            elif slide_type == 'team':
                deck.add_team_slide(
                    title=slide_data.get('title', ''),
                    members=slide_data.get('members', []),
                    slide_number=slide_number
                )
            
            elif slide_type == 'quote':
                deck.add_quote_slide(
                    quote=slide_data.get('quote', ''),
                    author=slide_data.get('author', ''),
                    role=slide_data.get('role', ''),
                    slide_number=slide_number
                )
            
            elif slide_type == 'stats':
                deck.add_stats_slide(
                    title=slide_data.get('title', ''),
                    stats=slide_data.get('stats', []),
                    slide_number=slide_number
                )
            
            elif slide_type == 'contact_info':
                deck.add_contact_info_slide(
                    title=slide_data.get('title', ''),
                    image_url=slide_data.get('image_url', ''),
                    contact_details=slide_data.get('contact_details', []),
                    slide_number=slide_number
                )
            
            elif slide_type == 'image_text_split':
                deck.add_image_text_split_slide(
                    title=slide_data.get('title', ''),
                    image_url=slide_data.get('image_url', ''),
                    content=slide_data.get('content', {}),
                    image_position=slide_data.get('image_position', 'left'),
                    slide_number=slide_number
                )
            
            elif slide_type == 'chart':
                deck.add_chart_slide(
                    title=slide_data.get('title', ''),
                    chart_type=slide_data.get('chart_type', 'line'),
                    chart_data=slide_data.get('chart_data', {}),
                    slide_number=slide_number
                )
            
            slide_counter += 1
        
        # Save to memory instead of disk (better for serverless)
        pptx_io = io.BytesIO()
        deck.save(pptx_io)
        pptx_io.seek(0)
        
        # Return the file
        return send_file(
            pptx_io,
            as_attachment=True,
            download_name=f"{data.get('title', 'presentation').replace(' ', '_')}.pptx",
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
    
    except Exception as e:
        import traceback
        error_trace = traceback.format_exc()
        print(f"Error: {error_trace}")
        return jsonify({'error': str(e), 'trace': error_trace}), 500

# Also add the /generate-presentation endpoint for compatibility
@app.route('/generate-presentation', methods=['POST'])
def generate_presentation():
    """Alias for create-presentation endpoint"""
    return create_presentation()

@app.route('/health', methods=['GET'])
def health():
    """Health check endpoint"""
    return jsonify({'status': 'ok', 'service': 'presentation-api'})

@app.route('/', methods=['GET'])
def home():
    """Info endpoint"""
    return jsonify({
        'status': 'running',
        'endpoints': {
            '/create-presentation': 'POST - Create presentation (original endpoint)',
            '/generate-presentation': 'POST - Create presentation (alias)',
            '/health': 'GET - Health check'
        }
    })

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
