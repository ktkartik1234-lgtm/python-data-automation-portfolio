import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import datetime

def generate_geetmp3_workbook(filepath="GeetMP3_YouTube_Song_Management.xlsx"):
    wb = openpyxl.Workbook()
    # Remove default sheet
    default_sheet = wb.active

    # Professional Executive Color Palette
    C_NAVY_DARK = "0F172A"       # Slate 900
    C_NAVY_HEADER = "1E293B"     # Slate 800
    C_CRIMSON = "991B1B"         # Deep Crimson Red (GeetMP3 brand color)
    C_CRIMSON_LIGHT = "FEF2F2"   # Red tint
    C_BLUE = "1D4ED8"            # Professional Blue
    C_BLUE_LIGHT = "EFF6FF"      # Blue tint
    C_EMERALD = "047857"         # Green accent
    C_EMERALD_LIGHT = "ECFDF5"   # Green tint
    C_PURPLE = "6D28D9"          # Purple accent
    C_PURPLE_LIGHT = "F5F3FF"    # Purple tint
    C_AMBER = "B45309"           # Amber accent
    C_AMBER_LIGHT = "FFFBEB"     # Amber tint
    C_ZEBRA = "F8FAFC"           # Light slate zebra stripe
    C_BORDER_LIGHT = "E2E8F0"    # Light border
    C_BORDER_MED = "94A3B8"      # Medium border
    C_WHITE = "FFFFFF"
    C_TEXT_MUTED = "64748B"

    # Font styles
    font_family = "Segoe UI"
    font_title = Font(name=font_family, size=15, bold=True, color="FFFFFF")
    font_subtitle = Font(name=font_family, size=9, italic=True, color="E2E8F0")
    font_section = Font(name=font_family, size=11, bold=True, color=C_NAVY_DARK)
    font_header = Font(name=font_family, size=9, bold=True, color="FFFFFF")
    font_data = Font(name=font_family, size=9, color="1E293B")
    font_bold_data = Font(name=font_family, size=9, bold=True, color="1E293B")
    font_link = Font(name=font_family, size=9, color="1D4ED8", underline="single")
    font_muted = Font(name=font_family, size=8, color=C_TEXT_MUTED)

    # Fill patterns
    fill_header = PatternFill(start_color=C_NAVY_HEADER, end_color=C_NAVY_HEADER, fill_type="solid")
    fill_total = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    fill_zebra = PatternFill(start_color=C_ZEBRA, end_color=C_ZEBRA, fill_type="solid")
    fill_white = PatternFill(start_color=C_WHITE, end_color=C_WHITE, fill_type="solid")

    # Card Fills
    fill_card_blue = PatternFill(start_color=C_BLUE_LIGHT, end_color=C_BLUE_LIGHT, fill_type="solid")
    fill_card_red = PatternFill(start_color=C_CRIMSON_LIGHT, end_color=C_CRIMSON_LIGHT, fill_type="solid")
    fill_card_green = PatternFill(start_color=C_EMERALD_LIGHT, end_color=C_EMERALD_LIGHT, fill_type="solid")
    fill_card_purple = PatternFill(start_color=C_PURPLE_LIGHT, end_color=C_PURPLE_LIGHT, fill_type="solid")
    fill_card_amber = PatternFill(start_color=C_AMBER_LIGHT, end_color=C_AMBER_LIGHT, fill_type="solid")

    # Borders
    thin_border_side = Side(style="thin", color=C_BORDER_LIGHT)
    border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    border_total = Border(
        top=Side(style="thin", color=C_BORDER_MED),
        bottom=Side(style="double", color=C_NAVY_DARK),
        left=thin_border_side,
        right=thin_border_side
    )
    border_card = Border(
        left=Side(style="thin", color=C_BORDER_MED),
        right=Side(style="thin", color=C_BORDER_MED),
        top=Side(style="thin", color=C_BORDER_MED),
        bottom=Side(style="thin", color=C_BORDER_MED)
    )

    # Alignments
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    align_wrap_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # =========================================================================
    # 1. CREATE DASHBOARD SHEET
    # =========================================================================
    ws_dash = wb.create_sheet(title="Dashboard")
    ws_dash.views.sheetView[0].showGridLines = True

    # Title Banner (Row 2 & 3)
    ws_dash.merge_cells("B2:K2")
    ws_dash["B2"] = "GEETMP3 YOUTUBE NETWORK - CENTRAL SONG MANAGEMENT DASHBOARD"
    ws_dash["B2"].font = font_title
    ws_dash["B2"].fill = PatternFill(start_color="881337", end_color="881337", fill_type="solid") # Deep GeetMP3 crimson
    ws_dash["B2"].alignment = Alignment(horizontal="center", vertical="center")
    ws_dash.row_dimensions[2].height = 36

    ws_dash.merge_cells("B3:K3")
    ws_dash["B3"] = "Executive Multi-Channel Song Catalog, Streaming Performance Analytics & Metadata Central"
    ws_dash["B3"].font = font_subtitle
    ws_dash["B3"].fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    ws_dash["B3"].alignment = Alignment(horizontal="center", vertical="center")
    ws_dash.row_dimensions[3].height = 20

    # -------------------------------------------------------------------------
    # KPI SUMMARY CARDS (Row 5 to 7)
    # -------------------------------------------------------------------------
    # Card 1: Total Songs (Col B:C) -> Points to E16
    ws_dash.merge_cells("B5:C5")
    ws_dash["B5"] = "TOTAL SONGS (ALL CHANNELS)"
    ws_dash["B5"].font = Font(name=font_family, size=8, bold=True, color=C_BLUE)
    ws_dash["B5"].fill = fill_card_blue
    ws_dash["B5"].alignment = align_center

    ws_dash.merge_cells("B6:C7")
    ws_dash["B6"] = "=E16"
    ws_dash["B6"].font = Font(name=font_family, size=20, bold=True, color="1E3A8A")
    ws_dash["B6"].fill = fill_card_blue
    ws_dash["B6"].alignment = align_center
    ws_dash["B6"].number_format = "#,##0"

    # Card 2: Total YouTube Views (Col D:E) -> Points to F16
    ws_dash.merge_cells("D5:E5")
    ws_dash["D5"] = "TOTAL YOUTUBE VIEWS"
    ws_dash["D5"].font = Font(name=font_family, size=8, bold=True, color=C_CRIMSON)
    ws_dash["D5"].fill = fill_card_red
    ws_dash["D5"].alignment = align_center

    ws_dash.merge_cells("D6:E7")
    ws_dash["D6"] = "=F16"
    ws_dash["D6"].font = Font(name=font_family, size=20, bold=True, color="991B1B")
    ws_dash["D6"].fill = fill_card_red
    ws_dash["D6"].alignment = align_center
    ws_dash["D6"].number_format = "#,##0"

    # Card 3: Avg Views Per Track (Col F:G) -> Points to G16
    ws_dash.merge_cells("F5:G5")
    ws_dash["F5"] = "AVERAGE VIEWS / SONG"
    ws_dash["F5"].font = Font(name=font_family, size=8, bold=True, color=C_EMERALD)
    ws_dash["F5"].fill = fill_card_green
    ws_dash["F5"].alignment = align_center

    ws_dash.merge_cells("F6:G7")
    ws_dash["F6"] = "=G16"
    ws_dash["F6"].font = Font(name=font_family, size=20, bold=True, color="065F46")
    ws_dash["F6"].fill = fill_card_green
    ws_dash["F6"].alignment = align_center
    ws_dash["F6"].number_format = "#,##0"

    # Card 4: Channels Monitored (Col H:I)
    ws_dash.merge_cells("H5:I5")
    ws_dash["H5"] = "CHANNELS MONITORED"
    ws_dash["H5"].font = Font(name=font_family, size=8, bold=True, color=C_PURPLE)
    ws_dash["H5"].fill = fill_card_purple
    ws_dash["H5"].alignment = align_center

    ws_dash.merge_cells("H6:I7")
    ws_dash["H6"] = 4
    ws_dash["H6"].font = Font(name=font_family, size=20, bold=True, color="5B21B6")
    ws_dash["H6"].fill = fill_card_purple
    ws_dash["H6"].alignment = align_center
    ws_dash["H6"].number_format = "#,##0"

    # Card 5: Streaming Coverage Rate (Col J:K) -> Average of Platform Audit
    ws_dash.merge_cells("J5:K5")
    ws_dash["J5"] = "STREAMING LINK COVERAGE"
    ws_dash["J5"].font = Font(name=font_family, size=8, bold=True, color=C_AMBER)
    ws_dash["J5"].fill = fill_card_amber
    ws_dash["J5"].alignment = align_center

    ws_dash.merge_cells("J6:K7")
    ws_dash["J6"] = "=AVERAGE(E31:E36)"
    ws_dash["J6"].font = Font(name=font_family, size=20, bold=True, color="92400E")
    ws_dash["J6"].fill = fill_card_amber
    ws_dash["J6"].alignment = align_center
    ws_dash["J6"].number_format = "0.0%"

    for r in range(5, 8):
        for c in range(2, 12):
            ws_dash.cell(row=r, column=c).border = border_card

    # -------------------------------------------------------------------------
    # SECTION 1: CHANNEL PERFORMANCE & SUMMARY TABLE (Row 10 to 16)
    # -------------------------------------------------------------------------
    ws_dash["B10"] = "CHANNEL PERFORMANCE & SUMMARY BREAKDOWN"
    ws_dash["B10"].font = font_section

    channel_headers = [
        "Channel Name", "Channel Focus / Genre", "Dedicated Sheet Link",
        "Total Songs", "Total YouTube Views", "Average Views / Song", "% Channel View Share"
    ]
    for i, h in enumerate(channel_headers, start=2):
        cell = ws_dash.cell(row=11, column=i, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_cell
    ws_dash.row_dimensions[11].height = 24

    channels_data = [
        ("GeetMP3", "Official Mainstream / Flagship Pop", "#'GeetMP3'!A1"),
        ("GeetMP3 Punjabi", "Regional Punjabi Music & Desi Beats", "#'GeetMP3 Punjabi'!A1"),
        ("GeetMP3 Haryanvi", "Haryanvi Hits, Folk Rap & Dance", "#'GeetMP3 Haryanvi'!A1"),
        ("Geet MP3 Devotional", "Spiritual Bhajans, Aarti & Chants", "#'Geet MP3 Devotional'!A1")
    ]

    for idx, (ch_name, ch_desc, ch_link) in enumerate(channels_data, start=12):
        r_num = idx
        # Col B: Channel Name
        c_b = ws_dash.cell(row=r_num, column=2, value=ch_name)
        c_b.font = font_bold_data
        c_b.alignment = align_left
        c_b.border = border_cell

        # Col C: Focus / Genre
        c_c = ws_dash.cell(row=r_num, column=3, value=ch_desc)
        c_c.font = font_data
        c_c.alignment = align_left
        c_c.border = border_cell

        # Col D: Quick Link
        c_d = ws_dash.cell(row=r_num, column=4, value=f'=HYPERLINK("{ch_link}", "➡ Open {ch_name} Sheet")')
        c_d.font = font_link
        c_d.alignment = align_center
        c_d.border = border_cell

        # Col E: Total Songs (Formula: counts non-empty songs in column D)
        c_e = ws_dash.cell(row=r_num, column=5, value=f'=COUNTA(\'{ch_name}\'!D4:D5000)')
        c_e.font = font_bold_data
        c_e.alignment = align_right
        c_e.border = border_cell
        c_e.number_format = "#,##0"

        # Col F: Total YouTube Views (Formula: sum of column M)
        c_f = ws_dash.cell(row=r_num, column=6, value=f'=SUM(\'{ch_name}\'!M4:M5000)')
        c_f.font = font_bold_data
        c_f.alignment = align_right
        c_f.border = border_cell
        c_f.number_format = "#,##0"

        # Col G: Avg Views / Song (Formula: Col F / Col E)
        c_g = ws_dash.cell(row=r_num, column=7, value=f'=IFERROR(F{r_num}/E{r_num}, 0)')
        c_g.font = font_data
        c_g.alignment = align_right
        c_g.border = border_cell
        c_g.number_format = "#,##0"

        # Col H: % Channel View Share (Formula: Col F / Total Views $F$16)
        c_h = ws_dash.cell(row=r_num, column=8, value=f'=IFERROR(F{r_num}/$F$16, 0)')
        c_h.font = font_data
        c_h.alignment = align_right
        c_h.border = border_cell
        c_h.number_format = "0.0%"

        ws_dash.row_dimensions[r_num].height = 20

    # Total Row for Channel Table (Row 16)
    ws_dash.merge_cells("B16:D16")
    c_tot_label = ws_dash["B16"]
    c_tot_label.value = "TOTAL / NETWORK SUMMARY"
    c_tot_label.font = Font(name=font_family, size=9, bold=True, color=C_NAVY_DARK)
    c_tot_label.alignment = Alignment(horizontal="right", vertical="center")
    c_tot_label.fill = fill_total

    for c in [2, 3, 4]:
        ws_dash.cell(row=16, column=c).fill = fill_total
        ws_dash.cell(row=16, column=c).border = border_total

    # Songs Total (Cell E16)
    c_tot_songs = ws_dash.cell(row=16, column=5, value="=SUM(E12:E15)")
    c_tot_songs.font = font_bold_data
    c_tot_songs.alignment = align_right
    c_tot_songs.fill = fill_total
    c_tot_songs.border = border_total
    c_tot_songs.number_format = "#,##0"

    # Views Total (Cell F16)
    c_tot_views = ws_dash.cell(row=16, column=6, value="=SUM(F12:F15)")
    c_tot_views.font = font_bold_data
    c_tot_views.alignment = align_right
    c_tot_views.fill = fill_total
    c_tot_views.border = border_total
    c_tot_views.number_format = "#,##0"

    # Overall Avg Views (Cell G16)
    c_tot_avg = ws_dash.cell(row=16, column=7, value="=IFERROR(F16/E16, 0)")
    c_tot_avg.font = font_bold_data
    c_tot_avg.alignment = align_right
    c_tot_avg.fill = fill_total
    c_tot_avg.border = border_total
    c_tot_avg.number_format = "#,##0"

    # Share Total (Cell H16)
    c_tot_share = ws_dash.cell(row=16, column=8, value="=SUM(H12:H15)")
    c_tot_share.font = font_bold_data
    c_tot_share.alignment = align_right
    c_tot_share.fill = fill_total
    c_tot_share.border = border_total
    c_tot_share.number_format = "0.0%"
    ws_dash.row_dimensions[16].height = 22

    # -------------------------------------------------------------------------
    # SECTION 2 & 3: RELEASE TYPE SUMMARY & LANGUAGE DISTRIBUTION (Row 19 to 26)
    # -------------------------------------------------------------------------
    ws_dash["B19"] = "RELEASE TYPE SUMMARY"
    ws_dash["B19"].font = font_section

    ws_dash["G19"] = "LANGUAGE DISTRIBUTION SUMMARY"
    ws_dash["G19"].font = font_section

    # Headers Release Type (Cols B:E)
    r_headers = ["Release Type", "Track Count", "% Catalog Share", "Total Views"]
    for i, h in enumerate(r_headers, start=2):
        c = ws_dash.cell(row=20, column=i, value=h)
        c.font = font_header
        c.fill = fill_header
        c.alignment = align_center
        c.border = border_cell

    # Headers Language (Cols G:J)
    l_headers = ["Language", "Track Count", "% Catalog Share", "Total Views"]
    for i, h in enumerate(l_headers, start=7):
        c = ws_dash.cell(row=20, column=i, value=h)
        c.font = font_header
        c.fill = fill_header
        c.alignment = align_center
        c.border = border_cell
    ws_dash.row_dimensions[20].height = 22

    release_types = ["Single", "Album", "Movie", "EP", "Other"]
    for idx, r_type in enumerate(release_types, start=21):
        # Type Name (Col B)
        ws_dash.cell(row=idx, column=2, value=r_type).alignment = align_left
        ws_dash.cell(row=idx, column=2).font = font_bold_data
        ws_dash.cell(row=idx, column=2).border = border_cell
        
        # Track Count formula (Col C)
        f_count = (f'=COUNTIF(\'GeetMP3\'!$F$4:$F$5000, "{r_type}") + '
                   f'COUNTIF(\'GeetMP3 Punjabi\'!$F$4:$F$5000, "{r_type}") + '
                   f'COUNTIF(\'GeetMP3 Haryanvi\'!$F$4:$F$5000, "{r_type}") + '
                   f'COUNTIF(\'Geet MP3 Devotional\'!$F$4:$F$5000, "{r_type}")')
        c_cnt = ws_dash.cell(row=idx, column=3, value=f_count)
        c_cnt.alignment = align_right
        c_cnt.font = font_data
        c_cnt.border = border_cell
        c_cnt.number_format = "#,##0"

        # % Catalog Share (Col D: Track Count / Total Songs $E$16)
        c_sh = ws_dash.cell(row=idx, column=4, value=f'=IFERROR(C{idx}/$E$16, 0)')
        c_sh.alignment = align_right
        c_sh.font = font_data
        c_sh.border = border_cell
        c_sh.number_format = "0.0%"

        # Total Views for Type (Col E)
        f_v = (f'=SUMIF(\'GeetMP3\'!$F$4:$F$5000, "{r_type}", \'GeetMP3\'!$M$4:$M$5000) + '
               f'SUMIF(\'GeetMP3 Punjabi\'!$F$4:$F$5000, "{r_type}", \'GeetMP3 Punjabi\'!$M$4:$M$5000) + '
               f'SUMIF(\'GeetMP3 Haryanvi\'!$F$4:$F$5000, "{r_type}", \'GeetMP3 Haryanvi\'!$M$4:$M$5000) + '
               f'SUMIF(\'Geet MP3 Devotional\'!$F$4:$F$5000, "{r_type}", \'Geet MP3 Devotional\'!$M$4:$M$5000)')
        c_vw = ws_dash.cell(row=idx, column=5, value=f_v)
        c_vw.alignment = align_right
        c_vw.font = font_data
        c_vw.border = border_cell
        c_vw.number_format = "#,##0"
        ws_dash.row_dimensions[idx].height = 20

    # Total Row for Release Types (Row 26)
    ws_dash.cell(row=26, column=2, value="TOTAL").font = font_bold_data
    ws_dash.cell(row=26, column=2).fill = fill_total
    ws_dash.cell(row=26, column=2).border = border_total
    for col_idx, f_tot in [(3, "=SUM(C21:C25)"), (4, "=SUM(D21:D25)"), (5, "=SUM(E21:E25)")]:
        c = ws_dash.cell(row=26, column=col_idx, value=f_tot)
        c.font = font_bold_data
        c.alignment = align_right
        c.fill = fill_total
        c.border = border_total
        c.number_format = "0.0%" if col_idx == 4 else "#,##0"

    # Languages (Cols G:J)
    languages = ["Punjabi", "Haryanvi", "Hindi", "Devotional / Sanskrit", "Other"]
    for idx, lang in enumerate(languages, start=21):
        ws_dash.cell(row=idx, column=7, value=lang).alignment = align_left
        ws_dash.cell(row=idx, column=7).font = font_bold_data
        ws_dash.cell(row=idx, column=7).border = border_cell

        f_l_cnt = (f'=COUNTIF(\'GeetMP3\'!$L$4:$L$5000, "{lang}") + '
                   f'COUNTIF(\'GeetMP3 Punjabi\'!$L$4:$L$5000, "{lang}") + '
                   f'COUNTIF(\'GeetMP3 Haryanvi\'!$L$4:$L$5000, "{lang}") + '
                   f'COUNTIF(\'Geet MP3 Devotional\'!$L$4:$L$5000, "{lang}")')
        c_l_cnt = ws_dash.cell(row=idx, column=8, value=f_l_cnt)
        c_l_cnt.alignment = align_right
        c_l_cnt.font = font_data
        c_l_cnt.border = border_cell
        c_l_cnt.number_format = "#,##0"

        c_l_sh = ws_dash.cell(row=idx, column=9, value=f'=IFERROR(H{idx}/$E$16, 0)')
        c_l_sh.alignment = align_right
        c_l_sh.font = font_data
        c_l_sh.border = border_cell
        c_l_sh.number_format = "0.0%"

        f_l_v = (f'=SUMIF(\'GeetMP3\'!$L$4:$L$5000, "{lang}", \'GeetMP3\'!$M$4:$M$5000) + '
                 f'SUMIF(\'GeetMP3 Punjabi\'!$L$4:$L$5000, "{lang}", \'GeetMP3 Punjabi\'!$M$4:$M$5000) + '
                 f'SUMIF(\'GeetMP3 Haryanvi\'!$L$4:$L$5000, "{lang}", \'GeetMP3 Haryanvi\'!$M$4:$M$5000) + '
                 f'SUMIF(\'Geet MP3 Devotional\'!$L$4:$L$5000, "{lang}", \'Geet MP3 Devotional\'!$M$4:$M$5000)')
        c_l_vw = ws_dash.cell(row=idx, column=10, value=f_l_v)
        c_l_vw.alignment = align_right
        c_l_vw.font = font_data
        c_l_vw.border = border_cell
        c_l_vw.number_format = "#,##0"

    # Total Row for Languages (Row 26)
    ws_dash.cell(row=26, column=7, value="TOTAL").font = font_bold_data
    ws_dash.cell(row=26, column=7).fill = fill_total
    ws_dash.cell(row=26, column=7).border = border_total
    for col_idx, f_tot in [(8, "=SUM(H21:H25)"), (9, "=SUM(I21:I25)"), (10, "=SUM(J21:J25)")]:
        c = ws_dash.cell(row=26, column=col_idx, value=f_tot)
        c.font = font_bold_data
        c.alignment = align_right
        c.fill = fill_total
        c.border = border_total
        c.number_format = "0.0%" if col_idx == 9 else "#,##0"

    # -------------------------------------------------------------------------
    # SECTION 4 & 5: STREAMING LINK AUDIT & YEARLY RELEASE TRACKER (Row 29 to 36)
    # -------------------------------------------------------------------------
    ws_dash["B29"] = "STREAMING PLATFORM DISTRIBUTION AUDIT"
    ws_dash["B29"].font = font_section

    ws_dash["G29"] = "YEARLY RELEASE CADENCE & PERFORMANCE"
    ws_dash["G29"].font = font_section

    plat_headers = ["Platform Name", "Catalog Column", "Tracks Linked", "Catalog Coverage", "Operational Status"]
    for i, h in enumerate(plat_headers, start=2):
        c = ws_dash.cell(row=30, column=i, value=h)
        c.font = font_header
        c.fill = fill_header
        c.alignment = align_center
        c.border = border_cell

    year_headers = ["Release Year", "Songs Released", "% Catalog Share", "Total YouTube Views"]
    for i, h in enumerate(year_headers, start=7):
        c = ws_dash.cell(row=30, column=i, value=h)
        c.font = font_header
        c.fill = fill_header
        c.alignment = align_center
        c.border = border_cell
    ws_dash.row_dimensions[30].height = 22

    # Platforms list
    platforms = [
        ("YouTube Video Links", "Col O", "O"),
        ("Spotify Track Links", "Col P", "P"),
        ("Apple Music Links", "Col Q", "Q"),
        ("JioSaavn Links", "Col R", "R"),
        ("Gaana Links", "Col S", "S"),
        ("Instagram Audio Links", "Col T", "T")
    ]
    for idx, (p_name, col_desc, col_let) in enumerate(platforms, start=31):
        # Name
        ws_dash.cell(row=idx, column=2, value=p_name).font = font_bold_data
        ws_dash.cell(row=idx, column=2).border = border_cell
        # Column Reference
        ws_dash.cell(row=idx, column=3, value=col_desc).font = font_muted
        ws_dash.cell(row=idx, column=3).alignment = align_center
        ws_dash.cell(row=idx, column=3).border = border_cell

        # Tracks Linked (Formula)
        f_p_cnt = (f'=COUNTA(\'GeetMP3\'!${col_let}$4:${col_let}$5000) + '
                   f'COUNTA(\'GeetMP3 Punjabi\'!${col_let}$4:${col_let}$5000) + '
                   f'COUNTA(\'GeetMP3 Haryanvi\'!${col_let}$4:${col_let}$5000) + '
                   f'COUNTA(\'Geet MP3 Devotional\'!${col_let}$4:${col_let}$5000)')
        c_pc = ws_dash.cell(row=idx, column=4, value=f_p_cnt)
        c_pc.font = font_data
        c_pc.alignment = align_right
        c_pc.border = border_cell
        c_pc.number_format = "#,##0"

        # Catalog Coverage % (Cell E: Tracks Linked / Total Songs $E$16)
        c_psh = ws_dash.cell(row=idx, column=5, value=f'=IFERROR(D{idx}/$E$16, 0)')
        c_psh.font = font_bold_data
        c_psh.alignment = align_right
        c_psh.border = border_cell
        c_psh.number_format = "0.0%"

        # Operational Status (Cell F)
        c_stat = ws_dash.cell(row=idx, column=6, value=f'=IF(E{idx}>=0.9,"Complete / 90%+","Pending Uploads")')
        c_stat.font = font_bold_data
        c_stat.alignment = align_center
        c_stat.border = border_cell
        ws_dash.row_dimensions[idx].height = 20

    # Years
    years = [("2024", 2024), ("2023", 2023), ("2022", 2022), ("2021", 2021), ("2020 & Earlier", 2020)]
    for idx, (yr_lbl, yr_val) in enumerate(years, start=31):
        ws_dash.cell(row=idx, column=7, value=yr_lbl).font = font_bold_data
        ws_dash.cell(row=idx, column=7).alignment = align_left
        ws_dash.cell(row=idx, column=7).border = border_cell

        if yr_val > 2020:
            f_yr_c = (f'=COUNTIFS(\'GeetMP3\'!$C$4:$C$5000, ">={yr_val}-01-01", \'GeetMP3\'!$C$4:$C$5000, "<={yr_val}-12-31") + '
                      f'COUNTIFS(\'GeetMP3 Punjabi\'!$C$4:$C$5000, ">={yr_val}-01-01", \'GeetMP3 Punjabi\'!$C$4:$C$5000, "<={yr_val}-12-31") + '
                      f'COUNTIFS(\'GeetMP3 Haryanvi\'!$C$4:$C$5000, ">={yr_val}-01-01", \'GeetMP3 Haryanvi\'!$C$4:$C$5000, "<={yr_val}-12-31") + '
                      f'COUNTIFS(\'Geet MP3 Devotional\'!$C$4:$C$5000, ">={yr_val}-01-01", \'Geet MP3 Devotional\'!$C$4:$C$5000, "<={yr_val}-12-31")')
            
            f_yr_v = (f'=SUMIFS(\'GeetMP3\'!$M$4:$M$5000, \'GeetMP3\'!$C$4:$C$5000, ">={yr_val}-01-01", \'GeetMP3\'!$C$4:$C$5000, "<={yr_val}-12-31") + '
                      f'SUMIFS(\'GeetMP3 Punjabi\'!$M$4:$M$5000, \'GeetMP3 Punjabi\'!$C$4:$C$5000, ">={yr_val}-01-01", \'GeetMP3 Punjabi\'!$C$4:$C$5000, "<={yr_val}-12-31") + '
                      f'SUMIFS(\'GeetMP3 Haryanvi\'!$M$4:$M$5000, \'GeetMP3 Haryanvi\'!$C$4:$C$5000, ">={yr_val}-01-01", \'GeetMP3 Haryanvi\'!$C$4:$C$5000, "<={yr_val}-12-31") + '
                      f'SUMIFS(\'Geet MP3 Devotional\'!$M$4:$M$5000, \'Geet MP3 Devotional\'!$C$4:$C$5000, ">={yr_val}-01-01", \'Geet MP3 Devotional\'!$C$4:$C$5000, "<={yr_val}-12-31")')
        else:
            f_yr_c = (f'=COUNTIFS(\'GeetMP3\'!$C$4:$C$5000, "<=2020-12-31", \'GeetMP3\'!$C$4:$C$5000, ">1900-01-01") + '
                      f'COUNTIFS(\'GeetMP3 Punjabi\'!$C$4:$C$5000, "<=2020-12-31", \'GeetMP3 Punjabi\'!$C$4:$C$5000, ">1900-01-01") + '
                      f'COUNTIFS(\'GeetMP3 Haryanvi\'!$C$4:$C$5000, "<=2020-12-31", \'GeetMP3 Haryanvi\'!$C$4:$C$5000, ">1900-01-01") + '
                      f'COUNTIFS(\'Geet MP3 Devotional\'!$C$4:$C$5000, "<=2020-12-31", \'Geet MP3 Devotional\'!$C$4:$C$5000, ">1900-01-01")')
            
            f_yr_v = (f'=SUMIFS(\'GeetMP3\'!$M$4:$M$5000, \'GeetMP3\'!$C$4:$C$5000, "<=2020-12-31") + '
                      f'SUMIFS(\'GeetMP3 Punjabi\'!$M$4:$M$5000, \'GeetMP3 Punjabi\'!$C$4:$C$5000, "<=2020-12-31") + '
                      f'SUMIFS(\'GeetMP3 Haryanvi\'!$M$4:$M$5000, \'GeetMP3 Haryanvi\'!$C$4:$C$5000, "<=2020-12-31") + '
                      f'SUMIFS(\'Geet MP3 Devotional\'!$M$4:$M$5000, \'Geet MP3 Devotional\'!$C$4:$C$5000, "<=2020-12-31")')

        c_yc = ws_dash.cell(row=idx, column=8, value=f_yr_c)
        c_yc.font = font_data
        c_yc.alignment = align_right
        c_yc.border = border_cell
        c_yc.number_format = "#,##0"

        c_ys = ws_dash.cell(row=idx, column=9, value=f'=IFERROR(H{idx}/$E$16, 0)')
        c_ys.font = font_data
        c_ys.alignment = align_right
        c_ys.border = border_cell
        c_ys.number_format = "0.0%"

        c_yv = ws_dash.cell(row=idx, column=10, value=f_yr_v)
        c_yv.font = font_data
        c_yv.alignment = align_right
        c_yv.border = border_cell
        c_yv.number_format = "#,##0"
        ws_dash.row_dimensions[idx].height = 20

    # Total Row for Yearly Cadence (Row 36)
    ws_dash.cell(row=36, column=7, value="TOTAL").font = font_bold_data
    ws_dash.cell(row=36, column=7).fill = fill_total
    ws_dash.cell(row=36, column=7).border = border_total
    for col_idx, f_tot in [(8, "=SUM(H31:H35)"), (9, "=SUM(I31:I35)"), (10, "=SUM(J31:J35)")]:
        c = ws_dash.cell(row=36, column=col_idx, value=f_tot)
        c.font = font_bold_data
        c.alignment = align_right
        c.fill = fill_total
        c.border = border_total
        c.number_format = "0.0%" if col_idx == 9 else "#,##0"
    ws_dash.row_dimensions[36].height = 22

    # Column widths for Dashboard
    dash_col_widths = {
        'A': 3,
        'B': 24,
        'C': 34,
        'D': 24,
        'E': 16,
        'F': 22,
        'G': 22,
        'H': 16,
        'I': 16,
        'J': 22,
        'K': 16
    }
    for col, width in dash_col_widths.items():
        ws_dash.column_dimensions[col].width = width

    # =========================================================================
    # 2. CREATE 4 DEDICATED CHANNEL SHEETS
    # =========================================================================
    channel_configs = [
        {
            "name": "GeetMP3",
            "banner_sub": "Official GeetMP3 Channel | Mainstream Pop, Superhit Singles & Flagship Releases",
            "theme_accent": "881337", # Deep Wine Crimson
            "samples": [
                [1, "GeetMP3", datetime.date(2019, 8, 13), "Lehanga", "Lehanga", "Single", "Jass Manak", "Jass Manak", "Jass Manak", "Sharry Nexus", "Satti Dhillon", "Punjabi", 1680000000, "03:45", "https://youtube.com/watch?v=IEDEtZ4UVtE", "https://open.spotify.com/track/4iJyoBOLtH0GxOxHG2j", "https://music.apple.com/song/lehanga", "https://jiosaavn.com/song/lehanga", "https://gaana.com/song/lehanga", "https://instagram.com/reels/audio/lehanga", "All-time highest streamed track; blockbuster viral reel anthem"],
                [2, "GeetMP3", datetime.date(2018, 11, 12), "Prada", "Prada", "Single", "Jass Manak", "Jass Manak", "Jass Manak", "Ar Deep", "Satti Dhillon", "Punjabi", 840000000, "03:28", "https://youtube.com/watch?v=dZ0fwJojhrs", "https://open.spotify.com/track/prada", "https://music.apple.com/song/prada", "https://jiosaavn.com/song/prada", "https://gaana.com/song/prada", "https://instagram.com/reels/audio/prada", "Breakthrough debut hit of Jass Manak"],
                [3, "GeetMP3", datetime.date(2020, 2, 14), "Suit Punjabi", "Age 19", "Album", "Jass Manak", "Jass Manak", "Jass Manak", "Sharry Nexus", "Satti Dhillon", "Punjabi", 390000000, "03:15", "https://youtube.com/watch?v=suitpunjabi", "https://open.spotify.com/track/suitpunjabi", "https://music.apple.com/song/suitpunjabi", "https://jiosaavn.com/song/suitpunjabi", "https://gaana.com/song/suitpunjabi", "https://instagram.com/reels/audio/suitpunjabi", "Featured track on Age 19 album"],
                [4, "GeetMP3", datetime.date(2018, 1, 5), "Yaar Beli", "Yaar Beli", "Single", "Guri", "Avtar Dhaliwal", "Deep Jandu", "Deep Jandu", "Parmish Verma", "Punjabi", 295000000, "03:40", "https://youtube.com/watch?v=yaar_beli", "https://open.spotify.com/track/yaar_beli", "https://music.apple.com/song/yaar_beli", "https://jiosaavn.com/song/yaar_beli", "https://gaana.com/song/yaar_beli", "https://instagram.com/reels/audio/yaar_beli", "Foundational hit for Guri on GeetMP3 label"],
                [5, "GeetMP3", datetime.date(2021, 6, 20), "Chete Karda", "Bad Munda", "Album", "Jass Manak", "Jass Manak", "Jass Manak", "Game Changerz", "Satti Dhillon", "Punjabi", 185000000, "03:32", "https://youtube.com/watch?v=chete_karda", "https://open.spotify.com/track/chete_karda", "https://music.apple.com/song/chete_karda", "https://jiosaavn.com/song/chete_karda", "https://gaana.com/song/chete_karda", "https://instagram.com/reels/audio/chete_karda", "Lead romantic single from Bad Munda album"]
            ]
        },
        {
            "name": "GeetMP3 Punjabi",
            "banner_sub": "GeetMP3 Punjabi Official | Regional Punjabi Music, Desi Beats & Folk Productions",
            "theme_accent": "1E3A8A", # Royal Navy
            "samples": [
                [1, "GeetMP3 Punjabi", datetime.date(2022, 3, 10), "Kala Tikka", "Kala Tikka", "Single", "Guri", "Guri", "Guri", "Snappy", "Satti Dhillon", "Punjabi", 85000000, "03:12", "https://youtube.com/watch?v=kala_tikka", "https://open.spotify.com/track/kala_tikka", "https://music.apple.com/song/kala_tikka", "https://jiosaavn.com/song/kala_tikka", "https://gaana.com/song/kala_tikka", "https://instagram.com/reels/audio/kala_tikka", "High-engagement wedding season single"],
                [2, "GeetMP3 Punjabi", datetime.date(2021, 11, 25), "Girlfriend", "No Compete", "EP", "Jass Manak", "Jass Manak", "Jass Manak", "Satti Dhillon", "Satti Dhillon", "Punjabi", 110000000, "02:58", "https://youtube.com/watch?v=girlfriend", "https://open.spotify.com/track/girlfriend", "https://music.apple.com/song/girlfriend", "https://jiosaavn.com/song/girlfriend", "https://gaana.com/song/girlfriend", "https://instagram.com/reels/audio/girlfriend", "Short-form video hit across reels & shorts"],
                [3, "GeetMP3 Punjabi", datetime.date(2023, 1, 18), "Chann Sitare (Folk Mix)", "Heritage Vol 1", "EP", "Ammy Virk", "Harmanjeet", "Gurmeet Singh", "Gurmeet Singh", "Navjit Buttar", "Punjabi", 64000000, "04:10", "https://youtube.com/watch?v=chann_sitare", "https://open.spotify.com/track/chann_sitare", "https://music.apple.com/song/chann_sitare", "https://jiosaavn.com/song/chann_sitare", "https://gaana.com/song/chann_sitare", "https://instagram.com/reels/audio/chann_sitare", "Acoustic traditional folk rendition"],
                [4, "GeetMP3 Punjabi", datetime.date(2022, 9, 5), "Majhail", "Majhail", "Single", "AP Dhillon / Gurinder", "Shinda Kahlon", "Gminxr", "Gminxr", "Rahul Chahal", "Punjabi", 92000000, "03:22", "https://youtube.com/watch?v=majhail", "https://open.spotify.com/track/majhail", "https://music.apple.com/song/majhail", "https://jiosaavn.com/song/majhail", "https://gaana.com/song/majhail", "https://instagram.com/reels/audio/majhail", "Collaboration release on regional imprint"],
                [5, "GeetMP3 Punjabi", datetime.date(2020, 8, 14), "Viah", "Age 19", "Album", "Jass Manak", "Jass Manak", "Jass Manak", "Snappy", "Satti Dhillon", "Punjabi", 145000000, "03:30", "https://youtube.com/watch?v=viah", "https://open.spotify.com/track/viah", "https://music.apple.com/song/viah", "https://jiosaavn.com/song/viah", "https://gaana.com/song/viah", "https://instagram.com/reels/audio/viah", "Evergreen marriage playlist anthem"]
            ]
        },
        {
            "name": "GeetMP3 Haryanvi",
            "banner_sub": "GeetMP3 Haryanvi | Desi Haryanvi Ragni, Urban Beats, Dance Hits & Folk Rap",
            "theme_accent": "065F46", # Deep Emerald Green
            "samples": [
                [1, "GeetMP3 Haryanvi", datetime.date(2022, 5, 12), "Gypsy (Balam Thanedar)", "Gypsy", "Single", "GD Kaur", "Pranjal Dahiya", "Diler Kharkiya", "Diler Kharkiya", "Meet Bros", "Haryanvi", 540000000, "03:10", "https://youtube.com/watch?v=balam_thanedar", "https://open.spotify.com/track/gypsy", "https://music.apple.com/song/gypsy", "https://jiosaavn.com/song/gypsy", "https://gaana.com/song/gypsy", "https://instagram.com/reels/audio/gypsy", "Mega viral dance track across North India"],
                [2, "GeetMP3 Haryanvi", datetime.date(2020, 10, 11), "52 Gaj Ka Daman", "52 Gaj", "Single", "Renuka Panwar", "Mukesh Jaji", "Aman Jaji", "Aman Jaji", "Sahil Sandhu", "Haryanvi", 1450000000, "03:05", "https://youtube.com/watch?v=52_gaj", "https://open.spotify.com/track/52_gaj", "https://music.apple.com/song/52_gaj", "https://jiosaavn.com/song/52_gaj", "https://gaana.com/song/52_gaj", "https://instagram.com/reels/audio/52_gaj", "Record-breaking Haryanvi cultural anthem"],
                [3, "GeetMP3 Haryanvi", datetime.date(2021, 4, 18), "Moto", "Desi Hip Hop", "Single", "Diler Kharkiya", "Ajay Hooda", "Aman Jaji", "Aman Jaji", "Ameet Choudhary", "Haryanvi", 620000000, "03:15", "https://youtube.com/watch?v=moto", "https://open.spotify.com/track/moto", "https://music.apple.com/song/moto", "https://jiosaavn.com/song/moto", "https://gaana.com/song/moto", "https://instagram.com/reels/audio/moto", "Superhit club & wedding dance staple"],
                [4, "GeetMP3 Haryanvi", datetime.date(2023, 7, 22), "Kothe Chad Lalkaru", "Desi Beats", "Single", "Renuka Panwar", "Rakesh Majreya", "Aman Jaji", "Aman Jaji", "Sahil Sandhu", "Haryanvi", 78000000, "03:20", "https://youtube.com/watch?v=kothe_chad", "https://open.spotify.com/track/kothe_chad", "https://music.apple.com/song/kothe_chad", "https://jiosaavn.com/song/kothe_chad", "https://gaana.com/song/kothe_chad", "https://instagram.com/reels/audio/kothe_chad", "Popular regional summer single"],
                [5, "GeetMP3 Haryanvi", datetime.date(2022, 11, 30), "Gulab", "Gulab EP", "EP", "Khasa Aala Chahar", "Khasa Aala Chahar", "Bamboo Beats", "Bamboo Beats", "Garry Khatrao", "Haryanvi", 95000000, "03:00", "https://youtube.com/watch?v=gulab", "https://open.spotify.com/track/gulab", "https://music.apple.com/song/gulab", "https://jiosaavn.com/song/gulab", "https://gaana.com/song/gulab", "https://instagram.com/reels/audio/gulab", "Top-charting street rap production"]
            ]
        },
        {
            "name": "Geet MP3 Devotional",
            "banner_sub": "Geet MP3 Devotional | Spiritual Bhajans, Aarti, Kirtan, Chants & Shabad Gurbani",
            "theme_accent": "78350F", # Spiritual Deep Amber
            "samples": [
                [1, "Geet MP3 Devotional", datetime.date(2021, 1, 1), "Shri Hanuman Chalisa", "Sankat Mochan", "Album", "Hariharan", "Traditional", "Lalit Sen", "Lalit Sen", "Rameshwar Prasad", "Hindi", 420000000, "09:45", "https://youtube.com/watch?v=hanuman_chalisa", "https://open.spotify.com/track/hanuman_chalisa", "https://music.apple.com/song/hanuman_chalisa", "https://jiosaavn.com/song/hanuman_chalisa", "https://gaana.com/song/hanuman_chalisa", "https://instagram.com/reels/audio/hanuman_chalisa", "Morning devotional staple, evergreen daily streams"],
                [2, "Geet MP3 Devotional", datetime.date(2022, 8, 19), "Achyutam Keshavam", "Krishna Leela", "Single", "Vikram Hazra", "Traditional", "Vikram Hazra", "Vikram Hazra", "Brijesh Sharma", "Devotional / Sanskrit", 165000000, "06:12", "https://youtube.com/watch?v=achyutam_keshavam", "https://open.spotify.com/track/achyutam_keshavam", "https://music.apple.com/song/achyutam_keshavam", "https://jiosaavn.com/song/achyutam_keshavam", "https://gaana.com/song/achyutam_keshavam", "https://instagram.com/reels/audio/achyutam_keshavam", "Janmashtami viral devotional anthem"],
                [3, "Geet MP3 Devotional", datetime.date(2020, 11, 14), "Maha Mrityunjaya Mantra 108 Times", "Shiva Shambhu", "Single", "Rattan Mohan Sharma", "Traditional", "Pandit Jasraj", "Pandit Jasraj", "Dharmesh Soni", "Devotional / Sanskrit", 210000000, "28:30", "https://youtube.com/watch?v=maha_mrityunjaya", "https://open.spotify.com/track/maha_mrityunjaya", "https://music.apple.com/song/maha_mrityunjaya", "https://jiosaavn.com/song/maha_mrityunjaya", "https://gaana.com/song/maha_mrityunjaya", "https://instagram.com/reels/audio/maha_mrityunjaya", "Continuous meditative chanting recording"],
                [4, "Geet MP3 Devotional", datetime.date(2023, 3, 22), "Jai Ganesh Deva (Aarti)", "Aarti Sangrah", "EP", "Anuradha Paudwal", "Traditional", "Arun Paudwal", "Arun Paudwal", "Surinder Sodhi", "Hindi", 130000000, "04:55", "https://youtube.com/watch?v=jai_ganesh", "https://open.spotify.com/track/jai_ganesh", "https://music.apple.com/song/jai_ganesh", "https://jiosaavn.com/song/jai_ganesh", "https://gaana.com/song/jai_ganesh", "https://instagram.com/reels/audio/jai_ganesh", "Festive Ganesh Chaturthi essential"],
                [5, "Geet MP3 Devotional", datetime.date(2022, 10, 2), "Mata Rani Ke Bhajan (Bhor Bhai Din)", "Mata Ki Bhentein", "Album", "Narendra Chanchal", "Traditional", "Surinder Sodhi", "Surinder Sodhi", "Pawan Kumar", "Hindi", 88000000, "07:20", "https://youtube.com/watch?v=bhor_bhai", "https://open.spotify.com/track/bhor_bhai", "https://music.apple.com/song/bhor_bhai", "https://jiosaavn.com/song/bhor_bhai", "https://gaana.com/song/bhor_bhai", "https://instagram.com/reels/audio/bhor_bhai", "Navratri festive devotional favorite"]
            ]
        }
    ]

    # Required 21 columns
    channel_columns = [
        ("Sr. No.", 8, align_center),
        ("Label", 18, align_left),
        ("Release Date", 14, align_center),
        ("Song Name", 28, align_left),
        ("Album / Project", 24, align_left),
        ("Release Type", 16, align_center),
        ("Singer", 22, align_left),
        ("Lyricist", 20, align_left),
        ("Composer", 20, align_left),
        ("Music Director", 20, align_left),
        ("Video Director", 20, align_left),
        ("Language", 18, align_center),
        ("YouTube Views", 16, align_right),
        ("Video Duration", 14, align_center),
        ("YouTube Link", 28, align_left),
        ("Spotify", 28, align_left),
        ("Apple Music", 28, align_left),
        ("JioSaavn", 28, align_left),
        ("Gaana", 28, align_left),
        ("Instagram Audio Link", 28, align_left),
        ("Remarks", 36, align_wrap_left)
    ]

    for cfg in channel_configs:
        ch_name = cfg["name"]
        ws = wb.create_sheet(title=ch_name)
        ws.views.sheetView[0].showGridLines = True

        # Header Title Banner (Row 1)
        ws.merge_cells("A1:Q1")
        ws["A1"] = f"GEETMP3 - {ch_name.upper()} SONG CATALOG & METADATA MANAGEMENT"
        ws["A1"].font = Font(name=font_family, size=13, bold=True, color="FFFFFF")
        ws["A1"].fill = PatternFill(start_color=cfg["theme_accent"], end_color=cfg["theme_accent"], fill_type="solid")
        ws["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

        # Quick Return to Dashboard Button (Row 1, Cols R-U)
        ws.merge_cells("R1:U1")
        ws["R1"] = '=HYPERLINK("#\'Dashboard\'!A1", "⬅ Return to Executive Dashboard")'
        ws["R1"].font = Font(name=font_family, size=10, bold=True, color="FFFFFF", underline="single")
        ws["R1"].fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
        ws["R1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 28

        # Subtitle Info (Row 2)
        ws.merge_cells("A2:U2")
        ws["A2"] = f"{cfg['banner_sub']} | Real-time synchronized with central dashboard. Scalable for 10,000+ tracks."
        ws["A2"].font = Font(name=font_family, size=8, italic=True, color="475569")
        ws["A2"].fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
        ws["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[2].height = 18

        # Table Column Headers (Row 3)
        for col_idx, (col_name, col_w, col_align) in enumerate(channel_columns, start=1):
            cell = ws.cell(row=3, column=col_idx, value=col_name)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = align_center
            cell.border = border_cell
            ws.column_dimensions[get_column_letter(col_idx)].width = col_w
        ws.row_dimensions[3].height = 26

        # Populate Sample Rows (Row 4 to 8)
        current_row = 4
        for sample in cfg["samples"]:
            for col_idx, val in enumerate(sample, start=1):
                cell = ws.cell(row=current_row, column=col_idx, value=val)
                cell.font = font_data
                cell.border = border_cell
                cell.fill = fill_white if current_row % 2 == 0 else fill_zebra
                
                col_name, _, col_align = channel_columns[col_idx - 1]
                cell.alignment = col_align

                if col_name == "Release Date":
                    cell.number_format = "YYYY-MM-DD"
                elif col_name == "YouTube Views":
                    cell.number_format = "#,##0"
                    cell.font = font_bold_data
                elif "Link" in col_name or col_name in ["Spotify", "Apple Music", "JioSaavn", "Gaana"]:
                    cell.font = font_link

            ws.row_dimensions[current_row].height = 20
            current_row += 1

        # Pre-format entry rows up to row 30
        for r_blank in range(current_row, 31):
            # Dynamic Sr No
            c_sr = ws.cell(row=r_blank, column=1, value=f'=IF(D{r_blank}<>"",ROW()-3,"")')
            c_sr.font = font_data
            c_sr.alignment = align_center
            c_sr.border = border_cell
            c_sr.fill = fill_white if r_blank % 2 == 0 else fill_zebra

            # Default Label to Sheet Name if song name is entered
            c_lbl = ws.cell(row=r_blank, column=2, value=f'=IF(D{r_blank}<>"", "{ch_name}", "")')
            c_lbl.font = font_data
            c_lbl.alignment = align_left
            c_lbl.border = border_cell
            c_lbl.fill = fill_white if r_blank % 2 == 0 else fill_zebra

            for c_idx in range(3, 22):
                c = ws.cell(row=r_blank, column=c_idx)
                c.font = font_data
                c.border = border_cell
                c.fill = fill_white if r_blank % 2 == 0 else fill_zebra
                col_name, _, col_align = channel_columns[c_idx - 1]
                c.alignment = col_align
                if col_name == "Release Date":
                    c.number_format = "YYYY-MM-DD"
                elif col_name == "YouTube Views":
                    c.number_format = "#,##0"
            ws.row_dimensions[r_blank].height = 20

        # Enable Auto-Filter on row 3 across all 21 columns
        ws.auto_filter.ref = f"A3:U{max(30, current_row - 1)}"

        # Freeze Panes at E4:
        # Columns A (Sr), B (Label), C (Date), D (Song Name) stay sticky when scrolling horizontally.
        # Rows 1-3 stay sticky when scrolling vertically.
        ws.freeze_panes = "E4"

        # Data Validation: Release Type Dropdown
        dv_rel = DataValidation(
            type="list",
            formula1='"Single,Album,Movie,EP,Other"',
            allow_blank=True,
            showDropDown=False
        )
        dv_rel.error = 'Please select a valid Release Type (Single, Album, Movie, EP, Other).'
        dv_rel.errorTitle = 'Invalid Release Type'
        dv_rel.prompt = 'Choose release type from dropdown'
        dv_rel.promptTitle = 'Release Type'
        ws.add_data_validation(dv_rel)
        dv_rel.add("F4:F5000")

        # Data Validation: Language Dropdown
        dv_lang = DataValidation(
            type="list",
            formula1='"Punjabi,Haryanvi,Hindi,Devotional / Sanskrit,Rajasthani,Other"',
            allow_blank=True,
            showDropDown=False
        )
        dv_lang.error = 'Please choose a valid language from the list.'
        dv_lang.errorTitle = 'Invalid Language'
        dv_lang.prompt = 'Select language from dropdown'
        dv_lang.promptTitle = 'Language'
        ws.add_data_validation(dv_lang)
        dv_lang.add("L4:L5000")

    # Remove default sheet if present
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])

    wb.save(filepath)
    print(f"Successfully generated: {filepath}")

if __name__ == "__main__":
    generate_geetmp3_workbook()
