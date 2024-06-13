import re

import requests
from bs4 import BeautifulSoup
from pptx import Presentation
from tqdm import tqdm

PRINT_ALL = True

superscript_lut = dict(zip((map(str, range(10))), "⁰¹²³⁴⁵⁶⁷⁸⁹"))  # From https://lingojam.com/TinyTextGenerator


def grab_psalm(psalm):
    # URL of the website to scrape
    url = f"https://bibel.github.io/EUe/ot/Ps_{psalm}.html"

    # Send a GET request to the URL
    response = requests.get(url)

    # Check if the request was successful
    if response.status_code == 200:
        # Parse the HTML content
        soup = BeautifulSoup(response.content, 'html.parser')

        # Find the element with class "biblehtmlcontent verses"
        content_div = soup.find('div', class_='biblehtmlcontent verses')

        if content_div:
            # Find all elements with class "v" within the content_div
            verses = content_div.find_all('div', class_='v')

            # Extract the verses in their individual lines and their number
            verse_objects = []
            for verse in verses:
                verse_text_raw = verse.get_text(strip=True)
                if not verse_text_raw[0].isdigit():
                    # This is the first verse -- it begins with the psalms title and purpose
                    continue
                # Extract and remove the leading verse number
                verse_number = re.findall(r"^\d+", verse_text_raw)[0]
                verse_text = re.sub(r"^\d+", "", verse_text_raw)
                # Remove trailing [sela], "/" or footnote number
                for _ in range(3):
                    # Repeat multiple times because footnote or sela could occur together
                    verse_text = re.sub(r"(/$)|(\d+$)|(\[Sela])$", "", verse_text).strip()
                verse_text = verse_text.replace("-", "—")

                verse_lines = [v.strip() for v in verse_text.split("/")]

                # At the first line of a verse, add the verse number in superscript
                verse_lines[0] = "".join(superscript_lut[num] for num in verse_number) + " " + verse_lines[0]
                verse_objects.append(verse_lines)

            # At the last line of the last verse, append a Halleluja
            # verse_objects[-1] += " — Halleluja!"
            # verse_objects.append(["— Halleluja!"])  # Append it as a standalone line
            return verse_objects
        else:
            print("The specified class 'biblehtmlcontent verses' was not found.")
            raise ValueError()
    else:
        print(f"Failed to retrieve the webpage. Status code: {response.status_code}")
        raise ConnectionError()


class PsalmWriter:
    top_margin = 0.05  # Inches

    def __init__(self, psalm_number: int, prs: Presentation, body_font_size, line_spacing_factor, space_after):
        self.psalm_number = psalm_number
        self.prs = prs
        self.body_font_size = body_font_size
        self.line_spacing_factor = line_spacing_factor
        self.space_after = space_after
        self.width_inch = prs.slide_width.inches
        self.height_inch = prs.slide_height.inches
        self.current_text_body_height = 0

        self.verses = grab_psalm(self.psalm_number)

    def add_slide(self, title_slide: bool):
        # Add a slide with a title bar and a body text element
        slide_layout = self.prs.slide_layouts[1 - int(title_slide)]
        return self.prs.slides.add_slide(slide_layout)

    @staticmethod
    def get_text_frame(slide, is_title_slide=False):
        # Get the body text element:
        # This is the second element in the title slide master, otherwise the first element
        text_frame = slide.placeholders[int(is_title_slide)].text_frame

        return text_frame, text_frame.paragraphs[0]

    @staticmethod
    def fill_title(slide, title):
        # Fill in the title text element
        title_text_frame = slide.placeholders[0].text_frame
        title_text_frame.text = title

    def point_to_inch(self, line_spacing_factor=1.0, after=0.0, point_to_pixel=1.33, dpi=96):
        pixels = (self.body_font_size * line_spacing_factor + after) * point_to_pixel
        return pixels / dpi

    def write_psalm(self):
        slide = self.add_slide(title_slide=True)
        self.fill_title(slide, f"Psalm {self.psalm_number}")
        text_frame, p = self.get_text_frame(slide, is_title_slide=True)
        self.current_text_body_height = 0

        for i, paragraph_lines in enumerate(self.verses):
            # Calculate if verse still fits on slide
            num_lines = len(paragraph_lines)
            paragraph_height = (num_lines - 1) * self.point_to_inch(line_spacing_factor=self.line_spacing_factor)
            paragraph_height += self.point_to_inch(line_spacing_factor=1.0, after=self.space_after)
            if text_frame._parent.top.inches + self.top_margin + self.current_text_body_height + paragraph_height > self.height_inch:
                # Add a slide with only a body text element
                slide = self.add_slide(title_slide=False)
                self.current_text_body_height = 0
                text_frame, p = self.get_text_frame(slide, is_title_slide=True)
            elif i > 0:
                p = text_frame.add_paragraph()

            # Join the lines with line breaks
            paragraph_text = '\n'.join(paragraph_lines)
            self.current_text_body_height += paragraph_height

            # Add text to paragraph
            run = p.add_run()
            run.text = paragraph_text
            # p.font.size = Pt(22)
            # p.font.color.rgb = RGBColor(255, 255, 255)  # White text color
            # p.font.name = 'Georgia'
            # p.space_after = Pt(14)  # Add space after each paragraph

            # Indent every second paragraph by 2 cm
            if i % 2 == 1:
                p.level = 1  # This will create an indent

            if i == len(self.verses) - 1:
                # At the last verse, append " - Halleluja" in italic
                halleluja_run = p.add_run()
                halleluja_run.font.italic = True
                halleluja_run.text = " — Halleluja!"

        # Save the presentation
        prs.save(f'Psalm_{self.psalm_number:03d}.pptx')


standard_spacing_font_specific = 1.65  # Depending on font!
spacing_factor = 1.2
line_spacing_factor = standard_spacing_font_specific * spacing_factor

if PRINT_ALL:
    for i in tqdm(list(range(1, 151))):
        try:
            prs = Presentation("Template.pptx")
            pw = PsalmWriter(i, prs, body_font_size=23, line_spacing_factor=line_spacing_factor, space_after=12)
            pw.write_psalm()
        except Exception as e:
            print(f"Error while writing psalm {i}:\n")
            raise e
else:
    prs = Presentation("Template.pptx")
    pw = PsalmWriter(84, prs, body_font_size=23, line_spacing_factor=line_spacing_factor, space_after=12)
    pw.write_psalm()
