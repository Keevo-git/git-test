from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_LINE_SPACING
from pathlib import Path

OUTPUT = Path("Day_1_IT_Fundamentals_Intensive_Reviewer.docx")

NAVY = "17365D"
BLUE = "2F5597"
LIGHT_BLUE = "D9EAF7"
PALE_BLUE = "EEF5FB"
GOLD = "D6A84B"
DARK = "222222"
GRAY = "666666"
WHITE = "FFFFFF"
GREEN = "2E7D32"
RED = "A61B1B"


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_margins(cell, top=100, start=100, bottom=100, end=100):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for m, v in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{m}"))
        if node is None:
            node = OxmlElement(f"w:{m}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(v))
        node.set(qn("w:type"), "dxa")


def set_repeat_table_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def set_cant_split(row):
    tr_pr = row._tr.get_or_add_trPr()
    cant_split = OxmlElement("w:cantSplit")
    tr_pr.append(cant_split)


def add_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = paragraph.add_run("Page ")
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor.from_string(GRAY)
    fld_char1 = OxmlElement("w:fldChar")
    fld_char1.set(qn("w:fldCharType"), "begin")
    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = "PAGE"
    fld_char2 = OxmlElement("w:fldChar")
    fld_char2.set(qn("w:fldCharType"), "end")
    run._r.append(fld_char1)
    run._r.append(instr_text)
    run._r.append(fld_char2)


def add_bottom_border(paragraph, color=BLUE, size=10, space=4):
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = p_pr.find(qn("w:pBdr"))
    if p_bdr is None:
        p_bdr = OxmlElement("w:pBdr")
        p_pr.append(p_bdr)
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size))
    bottom.set(qn("w:space"), str(space))
    bottom.set(qn("w:color"), color)
    p_bdr.append(bottom)


def keep_with_next(paragraph):
    p_pr = paragraph._p.get_or_add_pPr()
    keep = OxmlElement("w:keepNext")
    p_pr.append(keep)


def page_break_before(paragraph):
    p_pr = paragraph._p.get_or_add_pPr()
    node = OxmlElement("w:pageBreakBefore")
    p_pr.append(node)


def add_heading(doc, text, level=1, page_break=False):
    p = doc.add_paragraph(style=f"Heading {level}")
    p.add_run(text)
    if page_break:
        page_break_before(p)
    keep_with_next(p)
    return p


def add_label_paragraph(doc, label, text, style=None):
    p = doc.add_paragraph(style=style)
    r = p.add_run(label)
    r.bold = True
    r.font.color.rgb = RGBColor.from_string(BLUE)
    p.add_run(text)
    return p


def add_callout(doc, title, body, fill=PALE_BLUE, accent=BLUE):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    cell = table.cell(0, 0)
    set_cell_shading(cell, fill)
    set_cell_margins(cell, top=160, start=180, bottom=160, end=180)
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(title)
    r.bold = True
    r.font.color.rgb = RGBColor.from_string(accent)
    r.font.size = Pt(11)
    p2 = cell.add_paragraph(body)
    p2.paragraph_format.space_after = Pt(0)
    p2.paragraph_format.line_spacing = 1.05
    doc.add_paragraph().paragraph_format.space_after = Pt(0)
    return table


def add_table(doc, headers, rows, widths=None):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0]
    set_repeat_table_header(hdr)
    for i, h in enumerate(headers):
        cell = hdr.cells[i]
        set_cell_shading(cell, NAVY)
        set_cell_margins(cell, 100, 110, 100, 110)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(str(h))
        r.bold = True
        r.font.color.rgb = RGBColor.from_string(WHITE)
        r.font.size = Pt(9)
    for ri, row in enumerate(rows):
        cells = table.add_row().cells
        set_cant_split(table.rows[-1])
        for i, value in enumerate(row):
            cell = cells[i]
            set_cell_margins(cell, 80, 100, 80, 100)
            if ri % 2 == 1:
                set_cell_shading(cell, "F6F8FA")
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(str(value))
            r.font.size = Pt(8.8)
    if widths:
        for row in table.rows:
            for i, width in enumerate(widths):
                row.cells[i].width = Inches(width)
    return table


def add_bullet(doc, text, level=0):
    p = doc.add_paragraph(style="List Bullet" if level == 0 else "List Bullet 2")
    p.add_run(text)
    p.paragraph_format.space_after = Pt(2)
    return p


def add_numbered(doc, text):
    p = doc.add_paragraph(style="List Number")
    p.add_run(text)
    p.paragraph_format.space_after = Pt(3)
    return p


def add_question(doc, number, question, choices=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after = Pt(2)
    keep_with_next(p)
    r = p.add_run(f"{number}. ")
    r.bold = True
    r.font.color.rgb = RGBColor.from_string(NAVY)
    p.add_run(question)
    if choices:
        for letter, choice in zip("ABCD", choices):
            cp = doc.add_paragraph()
            cp.paragraph_format.left_indent = Inches(0.25)
            cp.paragraph_format.first_line_indent = Inches(-0.18)
            cp.paragraph_format.space_after = Pt(1)
            cp.add_run(f"{letter}. ").bold = True
            cp.add_run(choice)
    return p


def add_answer_item(doc, number, answer, explanation):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(f"{number}. {answer}")
    r.bold = True
    r.font.color.rgb = RGBColor.from_string(GREEN)
    p.add_run(f" — {explanation}")
    return p


def setup_styles(doc):
    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "Aptos"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Aptos")
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = RGBColor.from_string(DARK)
    normal.paragraph_format.space_after = Pt(5)
    normal.paragraph_format.line_spacing = 1.08

    for level, size, color in [(1, 18, NAVY), (2, 14, BLUE), (3, 11.5, NAVY)]:
        style = styles[f"Heading {level}"]
        style.font.name = "Aptos Display"
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "Aptos Display")
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor.from_string(color)
        style.paragraph_format.space_before = Pt(12 if level == 1 else 8)
        style.paragraph_format.space_after = Pt(5)
        style.paragraph_format.keep_with_next = True

    if "Question" not in styles:
        q = styles.add_style("Question", WD_STYLE_TYPE.PARAGRAPH)
        q.font.name = "Aptos"
        q.font.size = Pt(10.5)
        q.paragraph_format.space_before = Pt(5)
        q.paragraph_format.space_after = Pt(2)
        q.paragraph_format.keep_with_next = True


def create_document():
    doc = Document()
    setup_styles(doc)

    sec = doc.sections[0]
    sec.top_margin = Inches(0.65)
    sec.bottom_margin = Inches(0.65)
    sec.left_margin = Inches(0.72)
    sec.right_margin = Inches(0.72)
    sec.header_distance = Inches(0.28)
    sec.footer_distance = Inches(0.28)

    header = sec.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    hr = hp.add_run("BSIT LADDERIZED EXAM REVIEWER  •  DAY 1")
    hr.bold = True
    hr.font.size = Pt(8.5)
    hr.font.color.rgb = RGBColor.from_string(BLUE)
    add_bottom_border(hp, color=LIGHT_BLUE, size=6, space=3)

    footer = sec.footer
    fp = footer.paragraphs[0]
    fp.add_run("IT Fundamentals Intensive Review  |  ")
    fp.runs[0].font.size = Pt(8.5)
    fp.runs[0].font.color.rgb = RGBColor.from_string(GRAY)
    add_page_number(fp)

    # Cover page
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(80)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("BSIT LADDERIZED EXAMINATION")
    r.bold = True
    r.font.name = "Aptos Display"
    r.font.size = Pt(15)
    r.font.color.rgb = RGBColor.from_string(BLUE)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(20)
    r = p.add_run("7-DAY INTENSIVE REVIEWER")
    r.bold = True
    r.font.name = "Aptos Display"
    r.font.size = Pt(26)
    r.font.color.rgb = RGBColor.from_string(NAVY)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(12)
    r = p.add_run("DAY 1: IT FUNDAMENTALS")
    r.bold = True
    r.font.name = "Aptos Display"
    r.font.size = Pt(22)
    r.font.color.rgb = RGBColor.from_string(GOLD)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(16)
    r = p.add_run("Computing Concepts • Computer History • Number Systems\nCommunication Systems • IT Ethics • Philippine IT Laws")
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor.from_string(GRAY)

    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    cell = table.cell(0, 0)
    cell.width = Inches(5.7)
    set_cell_shading(cell, PALE_BLUE)
    set_cell_margins(cell, 180, 230, 180, 230)
    cp = cell.paragraphs[0]
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rr = cp.add_run("85 PRACTICE ITEMS WITH COMPLETE ANSWERS AND EXPLANATIONS")
    rr.bold = True
    rr.font.size = Pt(11)
    rr.font.color.rgb = RGBColor.from_string(NAVY)
    cp2 = cell.add_paragraph("50 Multiple Choice • 10 True/False • 10 Identification\n10 Number-System Exercises • 5 Situational Questions")
    cp2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp2.paragraph_format.space_after = Pt(0)

    doc.add_paragraph().paragraph_format.space_before = Pt(55)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Prepared as an intensive, exam-focused study companion")
    r.italic = True
    r.font.size = Pt(10)
    r.font.color.rgb = RGBColor.from_string(GRAY)

    doc.add_page_break()

    add_heading(doc, "How to Use This Reviewer", 1)
    add_callout(doc, "Recommended approach", "Study the core notes first, cover the answer key while answering the practice test, then review every incorrect item. Redo the number-system exercises without looking at the solutions.")

    add_heading(doc, "Day 1 Study Plan", 2)
    plan_rows = [
        ("1", "Computing Concepts", "60 minutes", "Explain hardware, software, IPO, CPU, memory, and computer types."),
        ("2", "History of Computing", "45 minutes", "Memorize major pioneers, inventions, and computer generations."),
        ("3", "Number Systems", "90 minutes", "Practice binary, octal, decimal, hexadecimal, and two's complement."),
        ("4", "Communication Systems", "45 minutes", "Identify communication elements, media, and transmission modes."),
        ("5", "Ethics and Philippine IT Laws", "60 minutes", "Differentiate privacy, security, cybercrime, and intellectual property."),
        ("6", "Practice Test and Error Review", "90 minutes", "Answer all items, score yourself, and revisit weak areas."),
    ]
    add_table(doc, ["Block", "Topic", "Time", "Target"], plan_rows, [0.45, 1.55, 0.9, 3.0])

    add_heading(doc, "Learning Objectives", 2)
    for item in [
        "Explain how a computer accepts data, processes it, stores it, and produces information.",
        "Distinguish hardware, software, firmware, memory, and storage components.",
        "Identify key milestones and the commonly taught generations of computers.",
        "Convert values among binary, octal, decimal, and hexadecimal systems.",
        "Recognize the elements, modes, media, and quality measures of data communication.",
        "Apply ethical principles and basic provisions of major Philippine IT-related laws to simple situations.",
    ]:
        add_bullet(doc, item)

    add_callout(doc, "Exam strategy", "When two options look correct, identify the exact term requested by the question. For number systems, write place values or group bits visibly to reduce careless errors.", fill="FFF7E5", accent="9A6A00")

    # Core notes
    add_heading(doc, "PART I — CORE REVIEW NOTES", 1, page_break=True)

    add_heading(doc, "1. Computing Concepts", 2)
    add_label_paragraph(doc, "Computer: ", "an electronic device that accepts data, follows stored instructions to process it, can store results, and produces information as output.")
    add_label_paragraph(doc, "Data: ", "raw facts, figures, symbols, measurements, or observations that have not yet been organized for a purpose.")
    add_label_paragraph(doc, "Information: ", "processed or organized data that has meaning and is useful for decision-making.")

    add_heading(doc, "The IPO Cycle", 3)
    ipo_rows = [
        ("Input", "Data and instructions enter the system.", "Keyboard, scanner, microphone, sensor"),
        ("Processing", "The CPU transforms data according to instructions.", "Arithmetic, comparison, sorting"),
        ("Output", "The system presents processed results.", "Monitor, printer, speaker"),
        ("Storage", "Data, programs, and results are retained for later use.", "SSD, HDD, flash drive, cloud storage"),
        ("Feedback", "Output is used to adjust the next input or process.", "Thermostat or automated control system"),
    ]
    add_table(doc, ["Stage", "Meaning", "Examples"], ipo_rows, [1.0, 3.0, 2.2])

    add_heading(doc, "Major Hardware Components", 3)
    hardware_rows = [
        ("Input devices", "Capture data or commands", "Keyboard, mouse, touchscreen, camera"),
        ("CPU", "Executes instructions and coordinates operations", "ALU, Control Unit, registers"),
        ("Primary memory", "Directly supports active processing", "RAM, ROM, cache"),
        ("Secondary storage", "Retains data long-term", "SSD, HDD, optical disc"),
        ("Output devices", "Present processed information", "Monitor, printer, speakers"),
        ("Communication devices", "Send and receive data", "NIC, modem, wireless adapter"),
    ]
    add_table(doc, ["Category", "Main Function", "Examples"], hardware_rows, [1.3, 2.8, 2.1])

    add_heading(doc, "Inside the CPU", 3)
    add_bullet(doc, "Arithmetic Logic Unit (ALU): performs arithmetic operations and logical comparisons.")
    add_bullet(doc, "Control Unit (CU): directs the fetch-decode-execute cycle and coordinates other components.")
    add_bullet(doc, "Registers: extremely fast, very small storage locations inside the CPU for current instructions, addresses, and intermediate values.")
    add_bullet(doc, "Cache: high-speed memory near or inside the CPU that stores frequently used data and instructions.")

    add_heading(doc, "Memory and Storage", 3)
    mem_rows = [
        ("RAM", "Volatile", "Read/write", "Holds active programs and data; contents are lost when power is removed."),
        ("ROM", "Nonvolatile", "Mostly read-only", "Stores permanent instructions such as firmware or startup routines."),
        ("Cache", "Volatile", "Very fast", "Reduces the time needed to access frequently used data."),
        ("SSD", "Nonvolatile", "Fast secondary storage", "Uses flash memory and has no moving mechanical parts."),
        ("HDD", "Nonvolatile", "Magnetic storage", "Stores large amounts of data using rotating platters."),
    ]
    add_table(doc, ["Component", "Volatility", "Access", "Purpose"], mem_rows, [0.8, 0.9, 1.3, 3.2])

    add_heading(doc, "Software Categories", 3)
    software_rows = [
        ("System software", "Manages hardware and provides a platform for applications.", "Operating system, device driver, utility"),
        ("Application software", "Helps users accomplish specific tasks.", "Word processor, browser, payroll system"),
        ("Programming software", "Supports the creation and testing of programs.", "Compiler, interpreter, IDE, debugger"),
        ("Firmware", "Software stored in nonvolatile memory and closely tied to hardware.", "Router firmware, BIOS/UEFI"),
    ]
    add_table(doc, ["Category", "Purpose", "Examples"], software_rows, [1.35, 3.0, 1.85])

    add_heading(doc, "Types of Computers", 3)
    types_rows = [
        ("Supercomputer", "Extremely complex, high-speed scientific calculations", "Climate modeling, simulations"),
        ("Mainframe", "Large-scale, highly reliable transaction processing", "Banking, airline reservations"),
        ("Server", "Provides services or resources to clients over a network", "Web, database, file server"),
        ("Workstation", "High-performance single-user professional computing", "Engineering, 3D design"),
        ("Personal computer", "General-purpose computing for one user", "Desktop, laptop"),
        ("Mobile device", "Portable, battery-powered computing", "Tablet, smartphone"),
        ("Embedded system", "Dedicated computing within a larger product", "Car controller, appliance, sensor"),
    ]
    add_table(doc, ["Type", "Primary Use", "Example"], types_rows, [1.35, 3.25, 1.6])

    add_callout(doc, "Remember", "RAM is primary, volatile memory. An SSD is secondary, nonvolatile storage. Faster does not automatically mean 'memory'; identify whether the component supports active processing or long-term retention.")

    add_heading(doc, "2. History and Generations of Computers", 2)
    history_rows = [
        ("Abacus", "Ancient counting device using movable beads."),
        ("Pascaline", "Blaise Pascal's mechanical calculator for addition and subtraction."),
        ("Stepped Reckoner", "Gottfried Wilhelm Leibniz's machine that extended mechanical calculation."),
        ("Jacquard loom", "Used punched cards to control weaving patterns; influenced programmable machines."),
        ("Difference Engine", "Charles Babbage's design for automatically computing mathematical tables."),
        ("Analytical Engine", "Babbage's general-purpose mechanical computer design with concepts resembling memory and processing."),
        ("Ada Lovelace", "Wrote an algorithm intended for the Analytical Engine; often called the first computer programmer."),
        ("Hollerith tabulator", "Used punched cards to process census data and accelerated large-scale data processing."),
    ]
    add_table(doc, ["Milestone / Person", "Significance"], history_rows, [1.8, 4.5])

    gen_rows = [
        ("First", "Vacuum tubes", "Machine language", "Large, hot, costly, power-hungry"),
        ("Second", "Transistors", "Assembly and early high-level languages", "Smaller, faster, more reliable"),
        ("Third", "Integrated circuits", "Operating systems and multiprogramming", "Greater reliability and reduced size"),
        ("Fourth", "Microprocessors", "Personal computers, GUIs, networking", "CPU functions integrated on a chip"),
        ("Fifth", "AI and parallel processing", "Natural-language and intelligent systems", "Commonly taught as an ongoing direction"),
    ]
    add_heading(doc, "Commonly Taught Computer Generations", 3)
    add_table(doc, ["Generation", "Key Technology", "Typical Development", "Exam Clue"], gen_rows, [0.75, 1.25, 2.5, 1.85])
    add_callout(doc, "Mnemonic", "Vacuum Tubes → Transistors → Integrated Circuits → Microprocessors → Artificial Intelligence. Focus first on the key technology that distinguishes each generation.", fill="FFF7E5", accent="9A6A00")

    add_heading(doc, "3. Number Systems", 2)
    add_label_paragraph(doc, "Radix or base: ", "the number of unique digits available in a positional number system.")
    num_rows = [
        ("Binary", "2", "0, 1", "101101₂"),
        ("Octal", "8", "0–7", "725₈"),
        ("Decimal", "10", "0–9", "156₁₀"),
        ("Hexadecimal", "16", "0–9, A–F", "3A7₁₆"),
    ]
    add_table(doc, ["System", "Base", "Valid Digits", "Example"], num_rows, [1.2, 0.75, 2.1, 1.2])

    add_heading(doc, "Core Conversion Methods", 3)
    add_numbered(doc, "Base n to decimal: multiply each digit by its positional weight and add the products.")
    add_numbered(doc, "Decimal to base n: repeatedly divide by the new base, record remainders, and read the remainders from bottom to top.")
    add_numbered(doc, "Binary to octal: group bits in sets of three from the radix point outward.")
    add_numbered(doc, "Binary to hexadecimal: group bits in sets of four from the radix point outward.")
    add_numbered(doc, "Octal or hexadecimal to binary: replace each digit with its 3-bit or 4-bit binary equivalent.")

    solved = [
        ("101101₂ → decimal", "32 + 8 + 4 + 1", "45₁₀"),
        ("11010110₂ → decimal", "128 + 64 + 16 + 4 + 2", "214₁₀"),
        ("156₁₀ → binary", "156 ÷ 2 repeatedly; read remainders upward", "10011100₂"),
        ("255₁₀ → hexadecimal", "255 ÷ 16 = 15 remainder 15; F and F", "FF₁₆"),
        ("3A7₁₆ → decimal", "3×16² + 10×16 + 7", "935₁₀"),
        ("725₈ → decimal", "7×8² + 2×8 + 5", "469₁₀"),
        ("1101011011₂ → hexadecimal", "0011 0110 1011 → 3 6 B", "36B₁₆"),
        ("2F₁₆ → binary", "2 → 0010; F → 1111", "00101111₂"),
        ("57₈ → binary", "5 → 101; 7 → 111", "101111₂"),
        ("−18 in 8-bit two's complement", "18 = 00010010; invert = 11101101; add 1", "11101110₂"),
    ]
    add_heading(doc, "Solved Examples", 3)
    add_table(doc, ["Problem", "Working", "Answer"], solved, [1.65, 3.3, 1.25])

    add_heading(doc, "Useful Bit Terms", 3)
    bit_rows = [
        ("Bit", "One binary digit: 0 or 1"),
        ("Nibble", "4 bits"),
        ("Byte", "8 bits"),
        ("Word", "A CPU-dependent group of bits processed as a unit, such as 32 or 64 bits"),
        ("Most significant bit", "The leftmost bit with the greatest positional value"),
        ("Least significant bit", "The rightmost bit with the smallest positional value"),
    ]
    add_table(doc, ["Term", "Meaning"], bit_rows, [1.7, 4.6])

    add_callout(doc, "Fast conversion check", "Binary-to-octal uses groups of 3 because 8 = 2³. Binary-to-hexadecimal uses groups of 4 because 16 = 2⁴.")

    add_heading(doc, "4. Elements of Computer and Communication Systems", 2)
    comm_rows = [
        ("Sender / source", "Originates the data or message."),
        ("Message", "The information being communicated."),
        ("Transmitter / encoder", "Converts the message into signals suitable for transmission."),
        ("Medium / channel", "The path through which signals travel."),
        ("Receiver / decoder", "Converts received signals into usable form."),
        ("Destination", "The intended person or device that receives the information."),
        ("Protocol", "A set of rules governing communication, format, timing, and error handling."),
        ("Noise", "Unwanted interference that may distort a signal."),
        ("Feedback", "A response indicating the result or status of communication."),
    ]
    add_table(doc, ["Element", "Role"], comm_rows, [1.7, 4.6])

    add_heading(doc, "Transmission Modes", 3)
    mode_rows = [
        ("Simplex", "One direction only", "Traditional radio or TV broadcast"),
        ("Half-duplex", "Both directions, but not at the same time", "Walkie-talkie"),
        ("Full-duplex", "Both directions simultaneously", "Telephone or video call"),
    ]
    add_table(doc, ["Mode", "Direction", "Example"], mode_rows, [1.25, 3.0, 2.05])

    add_heading(doc, "Transmission Media", 3)
    media_rows = [
        ("Twisted pair", "Guided", "Low cost; common in telephone and Ethernet cabling"),
        ("Coaxial cable", "Guided", "Shielded copper cable used in cable and broadband systems"),
        ("Fiber-optic cable", "Guided", "Uses light; high bandwidth and resistant to electromagnetic interference"),
        ("Radio", "Unguided", "Broadcast and wireless networking"),
        ("Microwave", "Unguided", "Directional line-of-sight and satellite communication"),
        ("Infrared", "Unguided", "Short-range line-of-sight communication"),
    ]
    add_table(doc, ["Medium", "Type", "Key Point"], media_rows, [1.45, 1.0, 3.85])

    add_heading(doc, "Communication Quality Measures", 3)
    add_bullet(doc, "Delivery: data reaches the correct destination.")
    add_bullet(doc, "Accuracy: data arrives without unwanted alteration or errors.")
    add_bullet(doc, "Timeliness: data arrives when it is needed.")
    add_bullet(doc, "Jitter: variation in packet delay; especially important for audio and video.")
    add_bullet(doc, "Bandwidth: theoretical capacity of a communication channel.")
    add_bullet(doc, "Throughput: actual useful data transferred per unit of time.")
    add_bullet(doc, "Latency: delay between sending data and receiving a response.")

    add_heading(doc, "5. IT Social, Ethical, and Professional Issues", 2)
    ethics_rows = [
        ("Ethics", "Principles used to judge right and responsible conduct."),
        ("Law", "Rules formally established and enforced by government."),
        ("Privacy", "A person's ability to control access to and use of personal information."),
        ("Confidentiality", "The duty to prevent unauthorized disclosure of entrusted information."),
        ("Security", "Safeguards that protect information and systems from threats."),
        ("Integrity", "Accuracy, completeness, and trustworthiness of information."),
        ("Availability", "Authorized users can access systems and information when needed."),
    ]
    add_table(doc, ["Concept", "Meaning"], ethics_rows, [1.5, 4.8])

    add_heading(doc, "Professional Responsibilities", 3)
    for item in [
        "Be honest about system capabilities, limitations, risks, qualifications, and results.",
        "Work only within your level of competence and seek supervision or training when necessary.",
        "Avoid harm and consider effects on users, organizations, and the public.",
        "Obtain authorization before accessing systems, accounts, devices, or data.",
        "Protect confidential information and use it only for legitimate purposes.",
        "Respect copyright, licenses, patents, trademarks, and other intellectual-property rights.",
        "Document work accurately and preserve appropriate audit trails.",
        "Disclose conflicts of interest and do not manipulate results for personal benefit.",
    ]:
        add_bullet(doc, item)

    add_heading(doc, "Major Philippine IT-Related Laws", 3)
    law_rows = [
        ("RA 10173", "Data Privacy Act of 2012", "Protects personal information in government and private information systems; establishes privacy principles, data-subject rights, and responsibilities of organizations processing personal data."),
        ("RA 10175", "Cybercrime Prevention Act of 2012", "Defines and penalizes offenses such as illegal access, illegal interception, data and system interference, misuse of devices, cybersquatting, computer-related fraud, forgery, and identity theft."),
        ("RA 8293", "Intellectual Property Code of the Philippines", "Provides the legal framework for copyright, trademarks, patents, and other intellectual-property rights."),
        ("RA 8792", "Electronic Commerce Act of 2000", "Recognizes electronic data messages, electronic documents, and electronic signatures, subject to legal requirements."),
    ]
    add_table(doc, ["Law", "Common Name", "Exam-Focused Purpose"], law_rows, [0.9, 1.65, 3.75])

    add_heading(doc, "Data Privacy Act: Key Ideas", 3)
    add_bullet(doc, "Transparency: the data subject should know the nature, purpose, and extent of processing.")
    add_bullet(doc, "Legitimate purpose: personal data must be processed for a lawful and declared purpose.")
    add_bullet(doc, "Proportionality: processing should be adequate, relevant, suitable, necessary, and not excessive.")
    add_bullet(doc, "Personal information controller (PIC): determines why and how personal data is processed.")
    add_bullet(doc, "Personal information processor (PIP): processes personal data on behalf of a PIC.")
    add_bullet(doc, "Data subjects have recognized rights, including rights involving information, access, correction, objection, erasure/blocking under applicable conditions, damages, and data portability.")

    add_heading(doc, "Cybercrime Terms to Distinguish", 3)
    cyber_rows = [
        ("Illegal access", "Accessing all or part of a computer system without right."),
        ("Illegal interception", "Intercepting nonpublic computer-data transmissions without right."),
        ("Data interference", "Damaging, deleting, deteriorating, altering, or suppressing computer data without right."),
        ("System interference", "Seriously hindering or interfering with a computer system's functioning."),
        ("Misuse of devices", "Producing, selling, procuring, importing, distributing, or possessing devices or credentials intended for cybercrime, under the law's conditions."),
        ("Cybersquatting", "Bad-faith acquisition of a domain name under circumstances described by law."),
        ("Computer-related identity theft", "Unauthorized acquisition, use, misuse, transfer, possession, alteration, or deletion of identifying information through ICT, under the law's elements."),
    ]
    add_table(doc, ["Term", "Simplified Meaning"], cyber_rows, [2.0, 4.3])

    add_heading(doc, "Intellectual Property Basics", 3)
    ip_rows = [
        ("Copyright", "Protects original literary and artistic works and their expression, subject to law."),
        ("Patent", "Protects a qualifying invention for a limited period in exchange for disclosure."),
        ("Trademark", "Identifies and distinguishes the source of goods or services."),
        ("Trade secret", "Valuable confidential business information protected through secrecy and related law or agreements."),
        ("Plagiarism", "Presenting another person's ideas or work as one's own without proper acknowledgment; an academic and ethical violation and may also involve copyright infringement."),
        ("Software license", "States how software may legally be installed, used, copied, modified, or distributed."),
    ]
    add_table(doc, ["Term", "Core Idea"], ip_rows, [1.45, 4.85])

    add_callout(doc, "Legal note", "This reviewer provides simplified academic summaries for exam preparation. Exact legal application depends on the complete statute, its implementing rules, later issuances, and the facts of a case.", fill="FFF2F2", accent=RED)

    # Quick sheets
    add_heading(doc, "PART II — QUICK MEMORY SHEETS", 1, page_break=True)
    add_heading(doc, "One-Minute Computer Generations", 2)
    add_table(doc, ["1st", "2nd", "3rd", "4th", "5th"], [["Vacuum tubes", "Transistors", "Integrated circuits", "Microprocessors", "AI / parallel processing"]], [1.2, 1.2, 1.2, 1.2, 1.2])

    add_heading(doc, "Binary–Octal–Hex Reference", 2)
    ref_rows = [
        ("0000", "0", "0"), ("0001", "1", "1"), ("0010", "2", "2"), ("0011", "3", "3"),
        ("0100", "4", "4"), ("0101", "5", "5"), ("0110", "6", "6"), ("0111", "7", "7"),
        ("1000", "—", "8"), ("1001", "—", "9"), ("1010", "—", "A"), ("1011", "—", "B"),
        ("1100", "—", "C"), ("1101", "—", "D"), ("1110", "—", "E"), ("1111", "—", "F"),
    ]
    add_table(doc, ["4-bit Binary", "Octal Digit*", "Hex Digit"], ref_rows, [2.0, 2.0, 2.0])
    p = doc.add_paragraph("*Octal conversion normally uses 3-bit groups: 000–111 correspond to 0–7.")
    p.runs[0].italic = True
    p.runs[0].font.size = Pt(9)

    add_heading(doc, "Laws at a Glance", 2)
    add_table(doc, ["Number", "Keyword", "Think of…"], [
        ("RA 10173", "Privacy", "Personal data, lawful processing, data-subject rights"),
        ("RA 10175", "Cybercrime", "Illegal access, interference, fraud, identity theft"),
        ("RA 8293", "Intellectual property", "Copyright, patent, trademark"),
        ("RA 8792", "Electronic commerce", "Electronic documents and signatures"),
    ], [1.1, 1.7, 3.5])

    add_heading(doc, "Commonly Confused Pairs", 2)
    confuse_rows = [
        ("Data vs Information", "Raw facts vs processed, meaningful results"),
        ("RAM vs Storage", "Temporary active workspace vs long-term retention"),
        ("Bandwidth vs Throughput", "Maximum capacity vs actual achieved transfer rate"),
        ("Privacy vs Security", "Control over personal information vs safeguards against threats"),
        ("Copyright vs Patent", "Original expression vs qualifying invention"),
        ("Simplex vs Half-duplex", "One direction only vs two directions at different times"),
        ("Illegal access vs Identity theft", "Unauthorized system access vs misuse of identifying information"),
    ]
    add_table(doc, ["Pair", "Difference"], confuse_rows, [2.0, 4.3])

    # Practice exam
    add_heading(doc, "PART III — PRACTICE TEST", 1, page_break=True)
    add_callout(doc, "Directions", "Answer all items before opening the answer-key section. Suggested time: 70–90 minutes. For multiple choice, select the best answer.")

    mcqs = [
        ("Which statement best defines a computer?", ["A device that only stores files", "An electronic device that accepts data, processes it according to instructions, stores results, and produces output", "A machine that always makes independent decisions", "Any appliance that uses electricity"], "B", "A computer performs input, processing, storage, and output under program control."),
        ("Raw facts and figures before processing are called:", ["Information", "Knowledge", "Data", "Output"], "C", "Data are unprocessed facts; information is processed and meaningful data."),
        ("Which unit performs arithmetic operations and logical comparisons?", ["Control Unit", "Arithmetic Logic Unit", "Power Supply Unit", "Network Interface Card"], "B", "The ALU handles arithmetic and logical operations."),
        ("Which component is volatile?", ["SSD", "ROM", "RAM", "Optical disc"], "C", "RAM loses its contents when electrical power is removed."),
        ("Which is system software?", ["Spreadsheet application", "Operating system", "Presentation file", "Digital photograph"], "B", "An operating system manages hardware and provides services for applications."),
        ("Firmware is best described as:", ["A printed user manual", "Software stored in nonvolatile memory and closely tied to hardware", "Any file downloaded from the internet", "A damaged computer part"], "B", "Firmware provides low-level control and is commonly stored in flash or ROM."),
        ("Which is an input device?", ["Printer", "Speaker", "Scanner", "Projector"], "C", "A scanner captures data and sends it into the computer."),
        ("What are registers?", ["Slow external storage devices", "Very small, high-speed storage locations inside the CPU", "Programs that protect against malware", "A type of printer"], "B", "Registers hold current instructions, addresses, and intermediate results."),
        ("A microcontroller inside a washing machine is an example of:", ["A mainframe", "An embedded system", "A supercomputer", "A workstation"], "B", "An embedded system performs a dedicated function inside a larger product."),
        ("Which computer type is designed for highly reliable, high-volume transaction processing?", ["Mainframe", "Tablet", "Game console", "Embedded sensor"], "A", "Mainframes commonly handle large-scale transactions for enterprises."),
        ("The first generation of electronic computers primarily used:", ["Transistors", "Integrated circuits", "Vacuum tubes", "Microprocessors"], "C", "Vacuum tubes are the defining technology of the first generation."),
        ("The defining hardware technology of the second generation was:", ["Transistors", "Vacuum tubes", "Microprocessors", "Quantum processors"], "A", "Transistors replaced vacuum tubes and improved size, speed, and reliability."),
        ("Integrated circuits are associated with which generation?", ["First", "Second", "Third", "Fifth"], "C", "The third generation is commonly identified by integrated circuits."),
        ("Microprocessors are the key technology of the:", ["First generation", "Second generation", "Fourth generation", "Pre-mechanical era"], "C", "The fourth generation integrated CPU functions onto microprocessor chips."),
        ("Who is often called the first computer programmer?", ["Ada Lovelace", "Herman Hollerith", "Blaise Pascal", "John von Neumann"], "A", "Ada Lovelace described an algorithm intended for Babbage's Analytical Engine."),
        ("Who designed the Analytical Engine?", ["Charles Babbage", "Gottfried Leibniz", "Joseph Jacquard", "Alan Turing"], "A", "Charles Babbage designed both the Difference Engine and Analytical Engine."),
        ("Herman Hollerith is strongly associated with:", ["Vacuum tubes", "Punched-card data processing", "The first web browser", "The transistor"], "B", "His tabulating system used punched cards to process census data."),
        ("The fifth generation is commonly associated with:", ["Mechanical gears only", "Artificial intelligence and parallel processing", "Vacuum tubes", "Punched cards only"], "B", "Many introductory texts describe AI-oriented systems as the fifth-generation direction."),
        ("What is the base of hexadecimal?", ["2", "8", "10", "16"], "D", "Hexadecimal uses sixteen symbols: 0–9 and A–F."),
        ("What is 101101₂ in decimal?", ["43", "44", "45", "46"], "C", "32 + 8 + 4 + 1 = 45."),
        ("What is 29₁₀ in binary?", ["11100₂", "11101₂", "11110₂", "11011₂"], "B", "29 = 16 + 8 + 4 + 1, so the bits are 11101."),
        ("What is 11111111₂ in hexadecimal?", ["EE₁₆", "EF₁₆", "FE₁₆", "FF₁₆"], "D", "Group as 1111 1111; each group equals F."),
        ("What is 7B₁₆ in decimal?", ["121", "122", "123", "124"], "C", "7×16 + 11 = 112 + 11 = 123."),
        ("What is 345₈ in decimal?", ["221", "229", "237", "245"], "B", "3×64 + 4×8 + 5 = 192 + 32 + 5 = 229."),
        ("Binary digits are grouped by how many when converting directly to octal?", ["2", "3", "4", "8"], "B", "One octal digit corresponds to three binary bits because 8 = 2³."),
        ("To obtain the two's complement of a binary number, you normally:", ["Reverse the order of the bits", "Add 2 to every bit", "Invert the bits and add 1", "Delete the most significant bit"], "C", "Two's complement is formed by taking the one's complement and adding one."),
        ("How many bits are in a nibble?", ["2", "4", "8", "16"], "B", "A nibble is four bits; a byte is eight bits."),
        ("Which is a character-encoding standard?", ["Unicode", "Ethernet", "Bluetooth", "SQL"], "A", "Unicode assigns code points to characters from many writing systems."),
        ("In data communication, the device or person that originates a message is the:", ["Noise", "Sender", "Medium", "Protocol"], "B", "The sender or source begins the communication process."),
        ("A protocol is:", ["A physical cable only", "A set of rules for communication", "A storage device", "A type of computer virus"], "B", "Protocols define matters such as format, timing, addressing, and error handling."),
        ("Which transmission mode permits simultaneous two-way communication?", ["Simplex", "Half-duplex", "Full-duplex", "Serial-only"], "C", "Full-duplex allows both sides to transmit at the same time."),
        ("Which medium uses pulses of light?", ["Coaxial cable", "Twisted pair", "Fiber-optic cable", "Microwave only"], "C", "Fiber-optic cables carry data using light and resist electromagnetic interference."),
        ("Unwanted interference that distorts a communication signal is called:", ["Feedback", "Noise", "Protocol", "Output"], "B", "Noise can alter or weaken transmitted signals."),
        ("The theoretical maximum capacity of a channel is its:", ["Latency", "Bandwidth", "Jitter", "Packet loss"], "B", "Bandwidth refers to capacity; throughput is the actual achieved transfer rate."),
        ("The delay between sending data and receiving a response is called:", ["Latency", "Accuracy", "Compression", "Encoding"], "A", "Latency measures communication delay."),
        ("Which Philippine law primarily protects personal information?", ["RA 8293", "RA 10173", "RA 8792", "RA 11032"], "B", "RA 10173 is the Data Privacy Act of 2012."),
        ("Which government body administers and implements the Data Privacy Act?", ["National Privacy Commission", "Commission on Audit", "Department of Tourism", "Professional Regulation Commission"], "A", "The National Privacy Commission is the country's privacy regulator under the Act."),
        ("Collecting only data that is necessary and not excessive reflects which principle?", ["Proportionality", "Permanence", "Publicity", "Anonymity"], "A", "Proportionality requires data processing to be adequate, relevant, necessary, and not excessive."),
        ("Which law defines and penalizes computer-related offenses such as illegal access?", ["RA 10175", "RA 8293", "RA 9003", "RA 8749"], "A", "RA 10175 is the Cybercrime Prevention Act of 2012."),
        ("Entering another person's computer account without permission most directly illustrates:", ["Legal backup", "Illegal access", "Fair use", "Data portability"], "B", "Unauthorized access to a computer system is the core idea of illegal access."),
        ("Bad-faith acquisition of a domain name under circumstances specified by law is called:", ["Packet switching", "Cybersquatting", "Encryption", "Data mining"], "B", "Cybersquatting is a cybercrime offense defined in RA 10175."),
        ("Which law is the Intellectual Property Code of the Philippines?", ["RA 10173", "RA 10175", "RA 8293", "RA 8792"], "C", "RA 8293 provides the principal intellectual-property framework."),
        ("Which right most directly protects an original novel or computer program as an original work?", ["Copyright", "Patent only", "Trademark only", "Domain registration"], "A", "Copyright protects original literary and artistic expression, including software as provided by law."),
        ("Which intellectual-property right identifies the source of goods or services?", ["Trademark", "Patent", "Copyright notice", "Data privacy consent"], "A", "A trademark distinguishes the source of goods or services."),
        ("Which intellectual-property right is most closely associated with a qualifying invention?", ["Copyright", "Patent", "Trademark", "Privacy right"], "B", "Patents protect qualifying inventions for a limited period, subject to legal requirements."),
        ("Presenting another person's work as your own without acknowledgment is:", ["Compression", "Plagiarism", "Encryption", "Authentication"], "B", "Plagiarism is an ethical and academic violation and can also overlap with copyright infringement."),
        ("A technician secretly opens a client's private files even though the repair does not require it. The clearest professional issue is:", ["Proper documentation", "Violation of authorization and confidentiality", "Efficient troubleshooting", "Data compression"], "B", "Access must be authorized and limited to a legitimate need."),
        ("A developer accepts a task requiring expertise they do not have and hides the risk from the client. Which duty is violated?", ["Professional competence and honesty", "Data portability", "Full-duplex communication", "Trademark registration"], "A", "Professionals should represent their skills honestly and work within their competence."),
        ("Which principle means preventing unauthorized disclosure of information?", ["Availability", "Confidentiality", "Redundancy", "Portability"], "B", "Confidentiality restricts information disclosure to authorized parties."),
        ("The CIA triad consists of:", ["Control, Internet, Access", "Confidentiality, Integrity, Availability", "Copyright, Invention, Attribution", "Communication, Input, Algorithm"], "B", "The CIA triad summarizes three core information-security objectives."),
    ]

    add_heading(doc, "A. Multiple Choice — 50 Items", 2)
    for idx, (q, choices, _, _) in enumerate(mcqs, 1):
        add_question(doc, idx, q, choices)

    tf_items = [
        ("RAM normally retains its contents after the computer is turned off.", False, "RAM is volatile and normally loses its contents without power."),
        ("A microprocessor is the defining hardware technology of the fourth computer generation.", True, "Fourth-generation systems are commonly associated with microprocessors."),
        ("One hexadecimal digit can represent exactly four binary bits.", True, "Sixteen possible hex values correspond to 2⁴ combinations."),
        ("Half-duplex communication allows both parties to transmit at exactly the same time.", False, "Half-duplex supports both directions, but only one direction at a time."),
        ("Fiber-optic cable is generally resistant to electromagnetic interference.", True, "It transmits light rather than electrical signals through copper."),
        ("Bandwidth and throughput always have exactly the same value.", False, "Bandwidth is theoretical capacity, while throughput is the actual achieved rate."),
        ("Ethical conduct can require more than merely obeying the minimum requirements of law.", True, "Law and ethics overlap, but ethical duties may be broader."),
        ("RA 10173 is known as the Cybercrime Prevention Act of 2012.", False, "RA 10173 is the Data Privacy Act; RA 10175 is the Cybercrime Prevention Act."),
        ("Copyright and plagiarism are identical concepts in every situation.", False, "Copyright is a legal right; plagiarism is an ethical or academic attribution violation, though they may overlap."),
        ("A software license can restrict how software is copied, installed, modified, or distributed.", True, "Licenses establish legally permitted uses of software."),
    ]
    add_heading(doc, "B. True or False — 10 Items", 2)
    for i, (statement, _, _) in enumerate(tf_items, 1):
        add_question(doc, i, statement)
        p = doc.add_paragraph("Answer: __________")
        p.paragraph_format.left_indent = Inches(0.25)
        p.runs[0].italic = True
        p.runs[0].font.color.rgb = RGBColor.from_string(GRAY)

    id_items = [
        ("The processed, meaningful result obtained from data.", "Information", "Information is data that has been organized or processed for use."),
        ("The CPU component that directs and coordinates operations.", "Control Unit", "The CU manages the fetch-decode-execute cycle and component coordination."),
        ("The high-speed memory that stores frequently used data near the CPU.", "Cache memory", "Cache reduces average access time for frequently used instructions and data."),
        ("The scientist and writer often called the first computer programmer.", "Ada Lovelace", "She described an algorithm intended for the Analytical Engine."),
        ("The base-16 number system.", "Hexadecimal", "Hexadecimal uses symbols 0–9 and A–F."),
        ("A group of eight bits.", "Byte", "Eight bits form one byte."),
        ("Two-way communication in which transmission occurs one direction at a time.", "Half-duplex", "Walkie-talkies are a common example."),
        ("Unwanted interference in a communication channel.", "Noise", "Noise may distort or degrade a signal."),
        ("The Philippine law known as the Data Privacy Act of 2012.", "Republic Act No. 10173", "RA 10173 governs personal-data protection in covered information systems."),
        ("An identifying sign that distinguishes the source of goods or services.", "Trademark", "Trademarks distinguish goods or services in commerce."),
    ]
    add_heading(doc, "C. Identification — 10 Items", 2)
    for i, (prompt, _, _) in enumerate(id_items, 1):
        add_question(doc, i, prompt)
        p = doc.add_paragraph("Answer: __________________________________________")
        p.paragraph_format.left_indent = Inches(0.25)
        p.runs[0].font.color.rgb = RGBColor.from_string(GRAY)

    conv_items = [
        ("Convert 10101101₂ to decimal.", "173₁₀", "128 + 32 + 8 + 4 + 1 = 173."),
        ("Convert 93₁₀ to binary.", "1011101₂", "93 = 64 + 16 + 8 + 4 + 1."),
        ("Convert 111010111₂ to hexadecimal.", "1D7₁₆", "Pad and group: 0001 1101 0111 → 1 D 7."),
        ("Convert A6₁₆ to binary.", "10100110₂", "A = 1010 and 6 = 0110."),
        ("Convert 5C₁₆ to decimal.", "92₁₀", "5×16 + 12 = 92."),
        ("Convert 671₈ to decimal.", "441₁₀", "6×64 + 7×8 + 1 = 384 + 56 + 1."),
        ("Convert 110110101₂ to octal.", "665₈", "Group as 110 110 101 → 6 6 5."),
        ("Convert 247₈ to binary.", "10100111₂", "2 = 010, 4 = 100, 7 = 111; remove the leading zero."),
        ("Convert 1023₁₀ to hexadecimal.", "3FF₁₆", "1023 = 3×256 + 15×16 + 15."),
        ("Represent −25 using 8-bit two's complement.", "11100111₂", "+25 = 00011001; invert = 11100110; add 1 = 11100111."),
    ]
    add_heading(doc, "D. Number-System Exercises — 10 Items", 2)
    for i, (prompt, _, _) in enumerate(conv_items, 1):
        add_question(doc, i, prompt)
        p = doc.add_paragraph("Solution / Answer: ______________________________________________")
        p.paragraph_format.left_indent = Inches(0.25)
        p.runs[0].font.color.rgb = RGBColor.from_string(GRAY)

    situational = [
        ("A school registration form asks for a student's religion, medical history, parents' income, and social-media passwords, although only name, contact details, and enrollment information are needed. Identify the most relevant privacy principle and state what should be changed.", "Proportionality / data minimization", "The school should collect only data that is adequate, relevant, necessary, and not excessive for the declared enrollment purpose. Unnecessary sensitive or unrelated information should not be required."),
        ("An intern finds an administrator's password written on a desk and uses it to view confidential payroll files out of curiosity. Identify the primary legal and ethical issues.", "Illegal or unauthorized access; confidentiality and authorization violations", "Using credentials without permission may constitute illegal access and clearly violates professional duties of authorization and confidentiality."),
        ("A developer copies a paid software library into a commercial project despite a license that forbids redistribution. What concept is primarily involved, and what should the developer do?", "Software licensing and intellectual property", "The developer must comply with the license, obtain a proper license, replace the library, or use a legally compatible alternative."),
        ("During a video call, sound arrives at irregular intervals even though the average connection speed is high. Which communication measure best describes the problem?", "Jitter", "Jitter is variation in packet delay and can disrupt real-time audio and video."),
        ("A company advertises that its new application is '100% secure' even though testing revealed unresolved vulnerabilities. Which professional principles are violated?", "Honesty, competence, risk disclosure, and avoidance of harm", "Professionals must not misrepresent security and should disclose material limitations and remediate risks."),
    ]
    add_heading(doc, "E. Situational Questions — 5 Items", 2)
    for i, (prompt, _, _) in enumerate(situational, 1):
        add_question(doc, i, prompt)
        for _ in range(3):
            p = doc.add_paragraph("________________________________________________________________________________")
            p.paragraph_format.space_after = Pt(1)
            p.runs[0].font.color.rgb = RGBColor.from_string("BBBBBB")

    # Answer key
    add_heading(doc, "PART IV — ANSWER KEY AND EXPLANATIONS", 1, page_break=True)
    add_callout(doc, "Scoring guide", "Count one point per item. A score of 72–85 suggests strong readiness; 60–71 is developing; below 60 means you should revisit the core notes and repeat the practice test. Conversion work should be checked for both method and final answer.", fill="EAF5EA", accent=GREEN)

    add_heading(doc, "A. Multiple Choice", 2)
    for idx, (_, _, answer, explanation) in enumerate(mcqs, 1):
        add_answer_item(doc, idx, answer, explanation)

    add_heading(doc, "B. True or False", 2)
    for idx, (_, truth, explanation) in enumerate(tf_items, 1):
        add_answer_item(doc, idx, "TRUE" if truth else "FALSE", explanation)

    add_heading(doc, "C. Identification", 2)
    for idx, (_, answer, explanation) in enumerate(id_items, 1):
        add_answer_item(doc, idx, answer, explanation)

    add_heading(doc, "D. Number-System Exercises", 2)
    for idx, (_, answer, explanation) in enumerate(conv_items, 1):
        add_answer_item(doc, idx, answer, explanation)

    add_heading(doc, "E. Situational Questions", 2)
    for idx, (_, answer, explanation) in enumerate(situational, 1):
        add_answer_item(doc, idx, answer, explanation)

    add_heading(doc, "Post-Test Error Log", 2)
    error_rows = [(str(i), "", "", "") for i in range(1, 11)]
    add_table(doc, ["No.", "Item / Topic Missed", "Why I Missed It", "Correct Rule or Method"], error_rows, [0.5, 1.7, 1.8, 2.3])

    add_heading(doc, "Last-Minute Checklist", 2)
    checklist = [
        "I can explain data, information, hardware, software, firmware, memory, and storage.",
        "I know ALU, Control Unit, registers, cache, RAM, ROM, SSD, and HDD.",
        "I can match each computer generation with its key technology.",
        "I can convert among binary, octal, decimal, and hexadecimal without a calculator.",
        "I remember that octal uses 3-bit groups and hexadecimal uses 4-bit groups.",
        "I can distinguish simplex, half-duplex, and full-duplex communication.",
        "I can distinguish bandwidth, throughput, latency, and jitter.",
        "I associate RA 10173 with privacy, RA 10175 with cybercrime, and RA 8293 with intellectual property.",
        "I can recognize ethical issues involving authorization, confidentiality, competence, honesty, and licensing.",
    ]
    for item in checklist:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(3)
        p.add_run("☐ ").font.color.rgb = RGBColor.from_string(BLUE)
        p.add_run(item)

    add_heading(doc, "Official References for the Law Summaries", 1, page_break=True)
    refs = [
        "Republic Act No. 10173 — Data Privacy Act of 2012. The Lawphil Project: https://lawphil.net/statutes/repacts/ra2012/ra_10173_2012.html",
        "Republic Act No. 10175 — Cybercrime Prevention Act of 2012. The Lawphil Project: https://lawphil.net/statutes/repacts/ra2012/ra_10175_2012.html",
        "Republic Act No. 8293 — Intellectual Property Code of the Philippines. The Lawphil Project: https://lawphil.net/statutes/repacts/ra1997/ra_8293_1997.html",
        "Republic Act No. 8792 — Electronic Commerce Act of 2000. The Lawphil Project: https://lawphil.net/statutes/repacts/ra2000/ra_8792_2000.html",
        "National Privacy Commission — Data Subject Rights: https://privacy.gov.ph/know-your-rights/",
    ]
    for ref in refs:
        add_bullet(doc, ref)

    add_callout(doc, "End of Day 1", "Before moving to Day 2, redo every incorrect question and complete the conversion exercises one more time without using the solved examples.", fill="EAF5EA", accent=GREEN)

    # Document metadata
    props = doc.core_properties
    props.title = "Day 1 IT Fundamentals Intensive Reviewer"
    props.subject = "BSIT Ladderized Examination — 7-Day Intensive Reviewer"
    props.author = "OpenAI ChatGPT"
    props.keywords = "IT fundamentals, reviewer, BSIT, number systems, ethics, Philippine IT laws"
    props.comments = "Exam-focused reviewer with 85 practice items and complete explanations."

    doc.save(OUTPUT)
    print(f"Created {OUTPUT.resolve()}")


if __name__ == "__main__":
    create_document()
