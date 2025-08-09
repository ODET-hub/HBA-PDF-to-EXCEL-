#!/usr/bin/env python3
"""
Script to create a test PDF with various content types for testing the PDF to Excel converter.
"""

from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch

def create_test_pdf():
    """Create a test PDF with various content types."""
    
    # Create the PDF document
    doc = SimpleDocTemplate("test_document.pdf", pagesize=letter)
    styles = getSampleStyleSheet()
    story = []
    
    # Title
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=24,
        spaceAfter=30,
        alignment=1  # Center alignment
    )
    story.append(Paragraph("Sample Business Report", title_style))
    story.append(Spacer(1, 20))
    
    # Header section
    header_style = ParagraphStyle(
        'CustomHeader',
        parent=styles['Heading2'],
        fontSize=16,
        spaceAfter=12,
        textColor=colors.darkblue
    )
    story.append(Paragraph("Executive Summary", header_style))
    
    # Paragraph content
    normal_style = styles['Normal']
    story.append(Paragraph(
        "This is a sample business report demonstrating various content types that can be "
        "extracted and converted to Excel format. The document contains tables, lists, "
        "headers, and paragraphs to test the intelligent content detection system.",
        normal_style
    ))
    story.append(Spacer(1, 20))
    
    # Another header
    story.append(Paragraph("Financial Data", header_style))
    
    # Table data
    table_data = [
        ['Month', 'Revenue', 'Expenses', 'Profit', 'Growth %'],
        ['January', '$125,000', '$85,000', '$40,000', '12.5%'],
        ['February', '$138,000', '$92,000', '$46,000', '15.0%'],
        ['March', '$142,000', '$88,000', '$54,000', '17.4%'],
        ['April', '$156,000', '$95,000', '$61,000', '13.0%'],
        ['May', '$168,000', '$102,000', '$66,000', '7.7%']
    ]
    
    # Create table
    table = Table(table_data, colWidths=[1.2*inch, 1.2*inch, 1.2*inch, 1.2*inch, 1.2*inch])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    story.append(table)
    story.append(Spacer(1, 20))
    
    # Another header
    story.append(Paragraph("Key Performance Indicators", header_style))
    
    # List content (as paragraphs with bullet points)
    list_items = [
        "• Customer satisfaction rate: 94.2%",
        "• Employee retention: 87%",
        "• Market share growth: 3.2%",
        "• Product quality score: 9.1/10",
        "• Operational efficiency: 92%"
    ]
    
    for item in list_items:
        story.append(Paragraph(item, normal_style))
    
    story.append(Spacer(1, 20))
    
    # Another header
    story.append(Paragraph("Department Performance", header_style))
    
    # Another table
    dept_data = [
        ['Department', 'Budget', 'Actual', 'Variance', 'Status'],
        ['Sales', '$500,000', '$485,000', '-$15,000', 'Under'],
        ['Marketing', '$200,000', '$210,000', '+$10,000', 'Over'],
        ['R&D', '$300,000', '$295,000', '-$5,000', 'Under'],
        ['Operations', '$150,000', '$148,000', '-$2,000', 'Under'],
        ['HR', '$100,000', '$102,000', '+$2,000', 'Over']
    ]
    
    dept_table = Table(dept_data, colWidths=[1.5*inch, 1.2*inch, 1.2*inch, 1.2*inch, 1.2*inch])
    dept_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 11),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.lightblue),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    story.append(dept_table)
    story.append(Spacer(1, 20))
    
    # Final paragraph
    story.append(Paragraph("Conclusion", header_style))
    story.append(Paragraph(
        "This sample document demonstrates the capability of the PDF to Excel converter "
        "to handle various content types including tables, lists, headers, and paragraphs. "
        "The converter should be able to intelligently detect and organize this content "
        "into separate Excel sheets for better analysis and reporting.",
        normal_style
    ))
    
    # Build the PDF
    doc.build(story)
    print("Test PDF created: test_document.pdf")

if __name__ == "__main__":
    create_test_pdf()
