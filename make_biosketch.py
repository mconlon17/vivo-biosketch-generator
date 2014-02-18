#!/usr/bin/env/python
"""
    make_biosketch.py -- Given a list of people involved with pilot
    awards, list their publications and grants in an RTF file suitable for
    use in Word

    Version 0.0 MC 2013-12-29
    --  Getting started with biosketch generation
    Version 0.1 MC 2014-01-02
    --  Basic working version for all sections
    Version 0.2 MC 2014-01-06
    --  Check for values for all attributes

    To Do
    --  add honors and service to positions and honors
    --  Month and year for degrees if available
    --  Location of educational instiution if available

"""

__author__ = "Michael Conlon"
__copyright__ = "Copyright 2014, University of Florida"
__license__ = "BSD 3-Clause license"
__version__ = "0.1"

import vivotools as vt
import datetime

# I wish rtf-ng was more organized and we didn't need all the imports below,
# but it isn't and we do.

from rtfng.Renderer import Renderer
from rtfng.Elements import Document, PAGE_NUMBER
from rtfng.Styles import TextStyle, ParagraphStyle
from rtfng.document.section import Section
from rtfng.document.paragraph import Paragraph, Table, Cell
from rtfng.document.character import B, I
from rtfng.PropertySets import MarginsPropertySet, BorderPropertySet, \
    FramePropertySet, TabPropertySet, TextPropertySet, ParagraphPropertySet
from rtfng.document.base import TAB

# Start here. Get person data.  Setup the document

print str(datetime.datetime.now())
person = vt.get_person("http://vivo.ufl.edu/individual/n25562",
    get_positions=True, get_publications=True, get_degrees=True,
    get_grants=True)
##person = vt.get_person("http://vivo.ufl.edu/individual/n31445",
##    get_positions=True, get_publications=True, get_degrees=True,
##    get_grants=True)
##person = vt.get_person("http://vivo.ufl.edu/individual/n11937",
##    get_positions=True, get_publications=True, get_degrees=True,
##    get_grants=True)
if 'overview' not in person:
    person['overview'] = ''
if 'era_commons' not in person:
    person['era_commons'] = ''
if 'first_name' not in person:
    person['first_name'] = ''
if 'last_name' not in person:
    person['last_name'] = ''
if 'preferred_title' not in person:
    person['preferred_title'] = ''
print str(datetime.datetime.now())

thin_edge = BorderPropertySet(width=11, style=BorderPropertySet.SINGLE)
topBottom = FramePropertySet(top=thin_edge, bottom=thin_edge)
bottom_frame = FramePropertySet(bottom=thin_edge)
bottom_right_frame = FramePropertySet(right=thin_edge, bottom=thin_edge)
right_frame = FramePropertySet(right=thin_edge)
top_frame = FramePropertySet(top=thin_edge)

doc = Document()
ss = doc.StyleSheet

# Set the margins for the section at 0.5 inch on all sides

ms = MarginsPropertySet(top=720, left=720, right=720, bottom=720)
section = Section(margins=ms)
doc.Sections.append(section)

# Improve the style sheet.  1440 twips to the inch

ps = ParagraphStyle('Title', TextStyle(TextPropertySet(ss.Fonts.Arial, 22,
    bold=True)).Copy(), ParagraphPropertySet(alignment=3, space_before=270,
    space_after=30))
ss.ParagraphStyles.append(ps)
ps = ParagraphStyle('Subtitle', TextStyle(TextPropertySet(ss.Fonts.Arial,
    16)).Copy(), ParagraphPropertySet(alignment=3, space_before=0,
    space_after=0))
ss.ParagraphStyles.append(ps)
ps = ParagraphStyle('SubtitleLeft', TextStyle(TextPropertySet(ss.Fonts.Arial,
    16)).Copy(), ParagraphPropertySet(space_before=0, space_after=0))
ss.ParagraphStyles.append(ps)
ps = ParagraphStyle('Heading 3', TextStyle(TextPropertySet(ss.Fonts.Arial, 22,
    bold=True)).Copy(), ParagraphPropertySet(space_before=180,
    space_after=60, tabs=[TabPropertySet(360)]))
ss.ParagraphStyles.append(ps)
ps = ParagraphStyle('Header', TextStyle(TextPropertySet(ss.Fonts.Arial,
    16)).Copy(), ParagraphPropertySet(left_indent=int((9.0/16.0)*1440),
    space_before=0, space_after=60))
ss.ParagraphStyles.append(ps)
ps = ParagraphStyle('Footer', TextStyle(TextPropertySet(ss.Fonts.Arial,
    16)).Copy(), ParagraphPropertySet(space_before=60, tabs=[\
    TabPropertySet(int(0.5*section.TwipsToRightMargin()),
    alignment=TabPropertySet.CENTER),
    TabPropertySet(int(0.5*section.TwipsToRightMargin()),
    alignment=TabPropertySet.RIGHT)]))
ss.ParagraphStyles.append(ps)

# Put in the header and footer

p = Paragraph(ss.ParagraphStyles.Header)
p.append("Program Director/Principal Investigator (Last, First, Middle): ")
section.Header.append(p)

p = Paragraph(ss.ParagraphStyles.Footer, top_frame)
p.append('PHS 398/2590 (Rev. 06/09)', TAB, 'Page ', PAGE_NUMBER, TAB,
    "Biographical Sketch Format Page")
section.Footer.append(p)

# Put in the top table

table = Table(5310, 270, 1170, 1440, 2610)

p1 = Paragraph(ss.ParagraphStyles.Title, "BIOGRAPHICAL SKETCH")
p2 = Paragraph(ss.ParagraphStyles.Subtitle)
p2.append('Provide the following information for the Senior/key personnel ' \
    'and other significant contributors in the order listed on Form Page 2.')
p3 = Paragraph(ss.ParagraphStyles.Subtitle)
p3.append("Follow this format for each person.  ",
    B("DO NOT EXCEED FOUR PAGES."))
c = Cell(p1, p2, p3, topBottom, span=5)
table.AddRow(c)

c = Cell(Paragraph(ss.ParagraphStyles.Subtitle, ' '), bottom_frame, span=5)
table.AddRow(c)

p1 = Paragraph(ss.ParagraphStyles.SubtitleLeft, 'NAME')
p2 = Paragraph(ss.ParagraphStyles.Normal, person['first_name'], ' ',
    person['last_name'])
c1 = Cell(p1, p2, bottom_right_frame, span=2)
c2 = Cell(Paragraph(ss.ParagraphStyles.SubtitleLeft, 'POSITION TITLE'), span=3)
table.AddRow(c1, c2)

p1 = Paragraph(ss.ParagraphStyles.SubtitleLeft,
    'eRA COMMONS USER NAME (credential, e.g., agency login)')
p2 = Paragraph(ss.ParagraphStyles.Normal, person['era_commons'])
c1 = Cell(p1, p2, bottom_right_frame, span=2)
c2 = Cell(Paragraph(ss.ParagraphStyles.Normal, person['preferred_title']),
    bottom_frame, span=3)
table.AddRow(c1, c2)

c = Cell(Paragraph(ss.ParagraphStyles.SubtitleLeft, "EDUCATION/TRAINING  ",
    I('(Begin with baccalaureate or other initial professional education,'
    ' such as nursing, include postdoctoral training and residency training'
    ' if applicable.)')), bottom_frame, span=5)
table.AddRow(c)

c1 = Cell(Paragraph(ss.ParagraphStyles.Subtitle, 'INSTITUTION AND LOCATION'),
    bottom_right_frame, alignment=Cell.ALIGN_CENTER)
p1 = Paragraph(ss.ParagraphStyles.Subtitle, 'DEGREE')
p2 = Paragraph(ss.ParagraphStyles.Subtitle, I('(if applicable)'))
c2 = Cell(p1, p2, bottom_right_frame, span=2)
c3 = Cell(Paragraph(ss.ParagraphStyles.Subtitle, 'MM/YY'),
    bottom_right_frame, alignment=Cell.ALIGN_CENTER)
c4 = Cell(Paragraph(ss.ParagraphStyles.Subtitle, 'FIELD OF STUDY'),
    bottom_frame, alignment=Cell.ALIGN_CENTER)
table.AddRow(c1, c2, c3, c4)

# The degrees

degrees = {}
for degree in person['degrees']:
    key = degree['end_date']['date']['year']+degree['major_field']
    degrees[key] = degree
last_degree = min(5, len(degrees))
ndegree = 0
for key in sorted(degrees.keys(), reverse=True):
    ndegree = ndegree + 1
    degree = degrees[key]
    if ndegree < last_degree:
        c1 = Cell(Paragraph(ss.ParagraphStyles.Normal, \
            degree['institution_name']), right_frame)
        c2 = Cell(Paragraph(ss.ParagraphStyles.Normal, \
            degree['degree_name']), right_frame, span=2)
        c3 = Cell(Paragraph(ss.ParagraphStyles.Normal, \
            degree['end_date']['date']['year']), right_frame)
        c4 = Cell(Paragraph(ss.ParagraphStyles.Normal, degree['major_field']))
        table.AddRow(c1, c2, c3, c4)
    else:
        c1 = Cell(Paragraph(ss.ParagraphStyles.Normal, \
            degree['institution_name']), bottom_right_frame)
        c2 = Cell(Paragraph(ss.ParagraphStyles.Normal, \
            degree['degree_name']), bottom_right_frame, span=2)
        c3 = Cell(Paragraph(ss.ParagraphStyles.Normal,
            degree['end_date']['date']['year']), bottom_right_frame)
        c4 = Cell(Paragraph(ss.ParagraphStyles.Normal, degree['major_field']), \
            bottom_frame)
        table.AddRow(c1, c2, c3, c4)
        break

section.append(table)

# Put in the note

p = Paragraph(ss.ParagraphStyles.Heading3)
p.append('NOTE: The Biographical Sketch may not exceed four pages.'
         ' Follow the formats and instructions below.')
section.append(p)

# Section A -- Personal Statemeent

p = Paragraph(ss.ParagraphStyles.Heading3)
p.append("A.", TAB, "Personal Statement")
section.append(p)
p = Paragraph(ss.ParagraphStyles.Normal)
p.append(person['overview'])
section.append(p)

# Section B -- Positions and Honors

p = Paragraph(ss.ParagraphStyles.Heading3)
p.append("B.", TAB, "Positions and Honors")
section.append(p)

positions = {}
for position in person['positions']:
    if 'start_date' in position and 'position_label' in position:
        key = position['start_date']['date']['year']+position['position_label']
        positions[key] = position
last_position = min(20, len(positions))
npos = 0
for key in sorted(positions.keys(), reverse=True):
    npos = npos + 1
    if npos > last_position:
        break
    position = positions[key]
    para_props = ParagraphPropertySet(tabs=[TabPropertySet(550),
        TabPropertySet(125), TabPropertySet(600)])
    para_props.SetFirstLineIndent(-1275)
    para_props.SetLeftIndent(1275)
    p = Paragraph(ss.ParagraphStyles.Normal, para_props)
    if 'end_date' in position:
        p.append(position['start_date']['date']['year'], TAB, '-', TAB,
            position['end_date']['date']['year'], TAB,
            position['position_label'], ', ', position['org_name'])
    else:
        p.append(position['start_date']['date']['year'], TAB, '-', TAB,
            TAB, position['position_label'], ', ', position['org_name'])
    section.append(p)

# Section C -- Selected Peer-reviewed Publications

p = Paragraph(ss.ParagraphStyles.Heading3)
p.append("C.", TAB, "Selected Peer-reviewed Publications")
section.append(p)

publications = {}
for pub in person['publications']:
    if 'date' in pub and 'publication_type' in pub and \
    pub['publication_type'] == 'academic-article':
        key = pub['date']['year']+pub['title']
        publications[key] = pub
last_pub = min(25, len(publications))
npub = 0
for key in sorted(publications.keys(), reverse=True):
    npub = npub + 1
    if npub > last_pub:
        break
    pub = publications[key]
    para_props = ParagraphPropertySet()
    para_props.SetFirstLineIndent(-720)
    para_props.SetLeftIndent(720)
    p = Paragraph(ss.ParagraphStyles.Normal, para_props)
    p.append(str(npub), ". ", vt.string_from_document(pub))
    section.append(p)

# Section D -- Research Support

p = Paragraph(ss.ParagraphStyles.Heading3)
p.append("D.", TAB, "Research Support")
section.append(p)

grants = {}
print person['grants']
for grant in person['grants']:
    if 'end_date' in grant and 'title' in grant and \
    datetime.datetime(int(grant['end_date']['date']['year']),
        int(grant['end_date']['date']['month']),
        int(grant['end_date']['date']['day'])) + \
        datetime.timedelta(days=3*365) > datetime.datetime.now():
        key = grant['end_date']['date']['year']+grant['title']
        grants[key] = grant
print grants
for key in sorted(grants.keys(), reverse=True):
    grant = grants[key]
    para_props = ParagraphPropertySet()
    para_props.SetFirstLineIndent(-720)
    para_props.SetLeftIndent(720)
    p = Paragraph(ss.ParagraphStyles.Normal, para_props)
    if 'role' in grant and grant['role'] == 'pi':
        grant_role = 'Principal Investigator'
    elif 'role' in grant and grant['role'] == 'coi':
        grant_role = 'Co-investigator'
    elif 'role' in grant and grant['role'] == 'inv':
        grant_role = 'Investigator'
    else:
        grant_role = ''
    p.append(grant['start_date']['datetime'][0:10], ' - ',
        grant['end_date']['datetime'][0:10], ', ', grant['title'], ', ',
        grant['awarded_by'], ', ', grant['sponsor_award_id'], ', ', grant_role)
    section.append(p)

# All Done.  Write the file

Renderer().Write(doc, file("biosketch.rtf", "w"))
print str(datetime.datetime.now())
