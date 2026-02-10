from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from corporate_template import CorporatePresentation
import json
import os
from datetime import datetime

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
                ]
            },
            {
                "type": "split_slide",
                "title": "Business Plan",
                "paragraphs": ["paragraph 1", "paragraph 2"]
            }
        ]
    }
    """
    try:
        data = request.json
        
        # Create presentation
        deck = CorporatePresentation()
        
        # Add title slide
        deck.add_title_slide(
            title=data.get('title', 'Presentation'),
            subtitle=data.get('subtitle', '')
        )
        
        # Add table of contents if sections provided
        if 'sections' in data and data['sections']:
            deck.add_table_of_contents(sections=data['sections'])
        
        # Add content slides
        for slide_data in data.get('slides', []):
            slide_type = slide_data.get('type')
            
            if slide_type == 'content_with_icons':
                deck.add_content_with_icons(
                    title=slide_data['title'],
                    items=slide_data['items']
                )
            
            elif slide_type == 'split_slide':
                deck.add_split_slide(
                    title=slide_data['title'],
                    paragraphs=slide_data['paragraphs']
                )
            
            elif slide_type == 'market_opportunities':
                deck.add_market_opportunities(
                    title=slide_data['title'],
                    items=slide_data['items']
                )
        
        # Save with timestamp to avoid conflicts
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"presentation_{timestamp}.pptx"
        filepath = os.path.join('presentations', filename)
        
        # Create presentations directory if it doesn't exist
        os.makedirs('presentations', exist_ok=True)
        
        deck.save(filepath)
        
        # Return the file
        return send_file(
            filepath,
            as_attachment=True,
            download_name=f"{data.get('title', 'presentation').replace(' ', '_')}.pptx",
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/health', methods=['GET'])
def health():
    """Health check endpoint"""
    return jsonify({'status': 'ok'})

if __name__ == '__main__':
    os.makedirs('presentations', exist_ok=True)
    app.run(host='0.0.0.0', port=5000, debug=True)