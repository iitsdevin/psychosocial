#!/usr/bin/env python3
"""
Build a macro-enabled .docm from the Level 1 PHIR Response .docx.
Creates a valid vbaProject.bin and injects it into the document ZIP.
"""

import os
import io
import struct
import zipfile
import shutil
from lxml import etree

UPLOAD_DIR = "/root/.claude/uploads/9ca4ee02-85b9-52bf-99ea-c1789f8f94da"
SOURCE_DOCX = os.path.join(UPLOAD_DIR, "13f36241-Level1PHIRResponse_9.DOCX")
OUTPUT_DOCM = "/home/user/psychosocial/Level1_PHIR_Response_WithMacro.docm"


# ---------- VBA Compression Algorithm (MS-OVBA 2.4.1) ----------

def compress_vba(data: bytes) -> bytes:
    """Compress data using the VBA compression algorithm.
    Uses uncompressed chunks for reliability."""
    if not data:
        return b'\x01'

    compressed = bytearray(b'\x01')  # Signature byte
    pos = 0

    while pos < len(data):
        chunk_data = data[pos:pos + 4096]
        chunk_len = len(chunk_data)
        pos += chunk_len

        if chunk_len == 4096:
            # Full uncompressed chunk: size field = 4095
            header = 4095 | (0b011 << 12) | (0 << 15)
        else:
            # Partial last chunk: pad to 4096 bytes
            padded = chunk_data + b'\x00' * (4096 - chunk_len)
            chunk_data = padded
            header = 4095 | (0b011 << 12) | (0 << 15)

        compressed.extend(struct.pack('<H', header))
        compressed.extend(chunk_data)

    return bytes(compressed)


def make_vba_module_stream(source_code: str) -> bytes:
    """Create a compressed VBA module stream."""
    source_bytes = source_code.encode('utf-8')
    return compress_vba(source_bytes)


def build_dir_stream(module_name: str) -> bytes:
    """Build the VBA dir stream (compressed)."""

    raw = bytearray()

    # --- PROJECTINFORMATION ---
    # PROJECTSYSKIND
    raw.extend(struct.pack('<HI', 0x0001, 4))
    raw.extend(struct.pack('<I', 0x00000001))  # Win32

    # PROJECTLCID
    raw.extend(struct.pack('<HI', 0x0002, 4))
    raw.extend(struct.pack('<I', 0x0409))

    # PROJECTLCIDINVOKE
    raw.extend(struct.pack('<HI', 0x0014, 4))
    raw.extend(struct.pack('<I', 0x0409))

    # PROJECTCODEPAGE
    raw.extend(struct.pack('<HI', 0x0003, 2))
    raw.extend(struct.pack('<H', 1252))

    # PROJECTNAME
    name = b"VBAProject"
    raw.extend(struct.pack('<HI', 0x0004, len(name)))
    raw.extend(name)

    # PROJECTDOCSTRING
    raw.extend(struct.pack('<HI', 0x0005, 0))
    raw.extend(struct.pack('<HI', 0x0040, 0))

    # PROJECTHELPFILEPATH
    raw.extend(struct.pack('<HI', 0x0006, 0))
    raw.extend(struct.pack('<HI', 0x003D, 0))

    # PROJECTHELPCONTEXT
    raw.extend(struct.pack('<HI', 0x0007, 4))
    raw.extend(struct.pack('<I', 0))

    # PROJECTLIBFLAGS
    raw.extend(struct.pack('<HI', 0x0008, 4))
    raw.extend(struct.pack('<I', 0))

    # PROJECTVERSION
    raw.extend(struct.pack('<HI', 0x0009, 4))
    raw.extend(struct.pack('<I', 0x61CC7004))

    # PROJECTCONSTANTS
    raw.extend(struct.pack('<HI', 0x000C, 0))
    raw.extend(struct.pack('<HI', 0x003C, 0))

    # --- PROJECTREFERENCES ---
    # Reference to stdole
    ref_name = b"stdole"
    raw.extend(struct.pack('<HI', 0x0016, len(ref_name)))
    raw.extend(ref_name)
    raw.extend(struct.pack('<HI', 0x003E, len(ref_name) * 2))
    raw.extend(ref_name.decode().encode('utf-16-le'))

    libid = b"*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\SysWOW64\\stdole2.tlb#OLE Automation"
    registered_data = bytearray()
    registered_data.extend(struct.pack('<I', len(libid)))
    registered_data.extend(libid)
    registered_data.extend(struct.pack('<II', 0, 0))
    raw.extend(struct.pack('<HI', 0x000D, len(registered_data)))
    raw.extend(registered_data)

    # --- PROJECTMODULES ---
    raw.extend(struct.pack('<HI', 0x000F, 2))
    raw.extend(struct.pack('<H', 2))  # 2 modules

    # PROJECTCOOKIE
    raw.extend(struct.pack('<HI', 0x0013, 2))
    raw.extend(struct.pack('<H', 0xFFFF))

    # --- Module: ThisDocument ---
    td_name = b"ThisDocument"
    raw.extend(struct.pack('<HI', 0x0019, len(td_name)))
    raw.extend(td_name)
    raw.extend(struct.pack('<HI', 0x0047, len(td_name) * 2))
    raw.extend(td_name.decode().encode('utf-16-le'))

    raw.extend(struct.pack('<HI', 0x001C, len(td_name)))
    raw.extend(td_name)
    raw.extend(struct.pack('<HI', 0x0032, len(td_name) * 2))
    raw.extend(td_name.decode().encode('utf-16-le'))

    raw.extend(struct.pack('<HI', 0x001E, 0))
    raw.extend(struct.pack('<HI', 0x0048, 0))

    raw.extend(struct.pack('<HI', 0x0031, 4))
    raw.extend(struct.pack('<I', 0))

    raw.extend(struct.pack('<HI', 0x001A, 4))
    raw.extend(struct.pack('<I', 0))

    raw.extend(struct.pack('<HI', 0x002C, 2))
    raw.extend(struct.pack('<H', 0xFFFF))

    raw.extend(struct.pack('<HI', 0x0021, 0))  # MODULETYPE = document
    raw.extend(struct.pack('<HI', 0x002B, 0))  # MODULE terminator

    # --- Module: PEEPOModule ---
    mod_name = module_name.encode('ascii')
    raw.extend(struct.pack('<HI', 0x0019, len(mod_name)))
    raw.extend(mod_name)
    raw.extend(struct.pack('<HI', 0x0047, len(mod_name) * 2))
    raw.extend(mod_name.decode().encode('utf-16-le'))

    raw.extend(struct.pack('<HI', 0x001C, len(mod_name)))
    raw.extend(mod_name)
    raw.extend(struct.pack('<HI', 0x0032, len(mod_name) * 2))
    raw.extend(mod_name.decode().encode('utf-16-le'))

    raw.extend(struct.pack('<HI', 0x001E, 0))
    raw.extend(struct.pack('<HI', 0x0048, 0))

    raw.extend(struct.pack('<HI', 0x0031, 4))
    raw.extend(struct.pack('<I', 0))

    raw.extend(struct.pack('<HI', 0x001A, 4))
    raw.extend(struct.pack('<I', 0))

    raw.extend(struct.pack('<HI', 0x002C, 2))
    raw.extend(struct.pack('<H', 0xFFFF))

    raw.extend(struct.pack('<HI', 0x0022, 0))  # MODULETYPE = procedural
    raw.extend(struct.pack('<HI', 0x002B, 0))  # MODULE terminator

    # End of dir stream
    raw.extend(struct.pack('<HI', 0x0010, 0))

    return compress_vba(bytes(raw))


def write_cfb(streams: dict) -> bytes:
    """Write a Compound File Binary (OLE) with the given streams."""
    SECTOR_SIZE = 512

    vba_streams = {}
    root_streams = {}
    for path, data in streams.items():
        if path.startswith('VBA/'):
            vba_streams[path[4:]] = data
        else:
            root_streams[path] = data

    entries = []

    # Entry 0: Root Entry
    entries.append({
        'name': 'Root Entry', 'type': 5, 'color': 1, 'data': b'',
    })

    # Entry 1: VBA storage
    entries.append({
        'name': 'VBA', 'type': 1, 'color': 1, 'data': b'',
    })

    # VBA sub-streams
    vba_stream_names = list(vba_streams.keys())
    for name in vba_stream_names:
        entries.append({
            'name': name, 'type': 2, 'color': 1, 'data': vba_streams[name],
        })

    # Root-level streams
    root_stream_names = list(root_streams.keys())
    for name in root_stream_names:
        entries.append({
            'name': name, 'type': 2, 'color': 1, 'data': root_streams[name],
        })

    num_entries = len(entries)

    # Set tree structure
    for e in entries:
        e['left'] = 0xFFFFFFFF
        e['right'] = 0xFFFFFFFF
        e['child'] = 0xFFFFFFFF

    # Root children: VBA + root streams
    root_children = [1] + list(range(2 + len(vba_stream_names), num_entries))
    _build_tree(entries, 0, root_children)

    # VBA children
    vba_children = list(range(2, 2 + len(vba_stream_names)))
    _build_tree(entries, 1, vba_children)

    # Allocate sectors
    all_stream_data = [(i, e['data']) for i, e in enumerate(entries)
                       if e['type'] == 2 and len(e['data']) > 0]

    entries_per_sector = SECTOR_SIZE // 128
    num_dir_sectors = (num_entries + entries_per_sector - 1) // entries_per_sector

    data_sector_map = {}
    total_data_sectors = 0
    for entry_idx, data in all_stream_data:
        num_sectors = max(1, (len(data) + SECTOR_SIZE - 1) // SECTOR_SIZE)
        data_sector_map[entry_idx] = (total_data_sectors, num_sectors)
        total_data_sectors += num_sectors

    num_fat_sectors = 1
    dir_start_sector = num_fat_sectors
    data_start_sector = dir_start_sector + num_dir_sectors
    total_sectors = data_start_sector + total_data_sectors

    while total_sectors > num_fat_sectors * (SECTOR_SIZE // 4):
        num_fat_sectors += 1
        dir_start_sector = num_fat_sectors
        data_start_sector = dir_start_sector + num_dir_sectors
        total_sectors = data_start_sector + total_data_sectors

    # Build FAT
    fat = [0xFFFFFFFD] * num_fat_sectors  # FAT sector markers

    for i in range(num_dir_sectors):
        fat.append(dir_start_sector + i + 1 if i < num_dir_sectors - 1 else 0xFFFFFFFE)

    for entry_idx, data in all_stream_data:
        start, count = data_sector_map[entry_idx]
        for j in range(count):
            fat.append(data_start_sector + start + j + 1 if j < count - 1 else 0xFFFFFFFE)

    fat_entries_per_sector = SECTOR_SIZE // 4
    while len(fat) < num_fat_sectors * fat_entries_per_sector:
        fat.append(0xFFFFFFFF)

    # Update entries
    for entry_idx, data in all_stream_data:
        start, count = data_sector_map[entry_idx]
        entries[entry_idx]['start_sector'] = data_start_sector + start
        entries[entry_idx]['size'] = len(data)

    entries[0]['start_sector'] = 0xFFFFFFFE
    entries[0]['size'] = 0
    entries[1]['start_sector'] = 0xFFFFFFFE
    entries[1]['size'] = 0

    # Build file
    output = bytearray()

    # Header
    header = bytearray(512)
    header[0:8] = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'
    struct.pack_into('<H', header, 24, 0x003E)
    struct.pack_into('<H', header, 26, 0x0003)
    struct.pack_into('<H', header, 28, 0xFFFE)
    struct.pack_into('<H', header, 30, 9)
    struct.pack_into('<H', header, 32, 6)
    struct.pack_into('<I', header, 40, 0)
    struct.pack_into('<I', header, 44, num_fat_sectors)
    struct.pack_into('<I', header, 48, dir_start_sector)
    struct.pack_into('<I', header, 56, 4096)
    struct.pack_into('<I', header, 60, 0xFFFFFFFE)
    struct.pack_into('<I', header, 64, 0)
    struct.pack_into('<I', header, 68, 0xFFFFFFFE)
    struct.pack_into('<I', header, 72, 0)
    for i in range(109):
        struct.pack_into('<I', header, 76 + i * 4,
                         i if i < num_fat_sectors else 0xFFFFFFFF)
    output.extend(header)

    # FAT sectors
    for s in range(num_fat_sectors):
        fat_sector = bytearray(SECTOR_SIZE)
        start = s * fat_entries_per_sector
        for i in range(fat_entries_per_sector):
            idx = start + i
            struct.pack_into('<I', fat_sector, i * 4,
                             fat[idx] if idx < len(fat) else 0xFFFFFFFF)
        output.extend(fat_sector)

    # Directory sectors
    dir_data = bytearray(num_dir_sectors * SECTOR_SIZE)
    for i, e in enumerate(entries):
        offset = i * 128
        name_utf16 = e['name'].encode('utf-16-le')
        name_len = len(name_utf16) + 2
        dir_data[offset:offset + min(len(name_utf16), 64)] = name_utf16[:64]
        struct.pack_into('<H', dir_data, offset + 64, min(name_len, 64))
        dir_data[offset + 66] = e['type']
        dir_data[offset + 67] = e['color']
        struct.pack_into('<I', dir_data, offset + 68, e['left'])
        struct.pack_into('<I', dir_data, offset + 72, e['right'])
        struct.pack_into('<I', dir_data, offset + 76, e['child'])
        struct.pack_into('<I', dir_data, offset + 116, e.get('start_sector', 0xFFFFFFFE))
        struct.pack_into('<I', dir_data, offset + 120, e.get('size', 0))
    output.extend(dir_data)

    # Data sectors
    for entry_idx, data in all_stream_data:
        start, count = data_sector_map[entry_idx]
        padded = data + b'\x00' * (count * SECTOR_SIZE - len(data))
        output.extend(padded)

    return bytes(output)


def _build_tree(entries, parent_idx, children):
    """Build a simple binary tree for directory entries."""
    if not children:
        return
    if len(children) == 1:
        entries[parent_idx]['child'] = children[0]
    elif len(children) == 2:
        entries[parent_idx]['child'] = children[0]
        entries[children[0]]['right'] = children[1]
    elif len(children) == 3:
        entries[parent_idx]['child'] = children[1]
        entries[children[1]]['left'] = children[0]
        entries[children[1]]['right'] = children[2]
    else:
        mid = len(children) // 2
        entries[parent_idx]['child'] = children[mid]
        # Left subtree
        for i in range(0, mid):
            if i < mid - 1:
                entries[children[i]]['right'] = children[i + 1]
        entries[children[mid]]['left'] = children[0]
        # Right subtree
        if mid + 1 < len(children):
            entries[children[mid]]['right'] = children[mid + 1]
            for i in range(mid + 1, len(children) - 1):
                entries[children[i]]['right'] = children[i + 1]


def inject_vba_into_docx(docx_path, docm_path, vba_bin):
    """Take a .docx, add vbaProject.bin, update content types, save as .docm"""
    temp_path = docm_path + '.tmp'

    with zipfile.ZipFile(docx_path, 'r') as zin:
        with zipfile.ZipFile(temp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)

                if item.filename == '[Content_Types].xml':
                    root = etree.fromstring(data)
                    ns = 'http://schemas.openxmlformats.org/package/2006/content-types'

                    # Add VBA content type
                    override = etree.SubElement(root, f'{{{ns}}}Override')
                    override.set('PartName', '/word/vbaProject.bin')
                    override.set('ContentType', 'application/vnd.ms-office.vbaProject')

                    # Change main content type to macroEnabled
                    for elem in root:
                        ct = elem.get('ContentType', '')
                        if 'wordprocessingml.document.main+xml' in ct:
                            elem.set('ContentType',
                                'application/vnd.ms-word.document.macroEnabled.main+xml')

                    data = etree.tostring(root, xml_declaration=True,
                                          encoding='UTF-8', standalone=True)

                elif item.filename == 'word/_rels/document.xml.rels':
                    root = etree.fromstring(data)
                    ns = 'http://schemas.openxmlformats.org/package/2006/relationships'

                    max_id = 0
                    for rel in root:
                        rid = rel.get('Id', '')
                        if rid.startswith('rId'):
                            try:
                                max_id = max(max_id, int(rid[3:]))
                            except ValueError:
                                pass

                    new_rel = etree.SubElement(root, f'{{{ns}}}Relationship')
                    new_rel.set('Id', f'rId{max_id + 1}')
                    new_rel.set('Type',
                        'http://schemas.microsoft.com/office/2006/relationships/vbaProject')
                    new_rel.set('Target', 'vbaProject.bin')

                    data = etree.tostring(root, xml_declaration=True,
                                          encoding='UTF-8', standalone=True)

                zout.writestr(item, data)

            # Add vbaProject.bin
            zout.writestr('word/vbaProject.bin', vba_bin)

    os.replace(temp_path, docm_path)


# ===== VBA CODE =====
MODULE_NAME = "PEEPOModule"

VBA_CODE = 'Attribute VB_Name = "PEEPOModule"\r\n' + r"""Option Explicit

' ---------------------------------------------------------------
'  PEEPO Work-Factor Picker
'  Run this macro (Alt+F8 > SelectPEEPOFactors) to select
'  psychosocial factors and insert them into the
'  "Outline any relevant work or non-work factors" cell.
' ---------------------------------------------------------------

Public Sub SelectPEEPOFactors()
    Dim cats(1 To 5) As String
    cats(1) = "People (Human Factors)"
    cats(2) = "Equipment (Resources & Equipment)"
    cats(3) = "Environment (Physical & Workplace)"
    cats(4) = "Processes (Tasks & Procedures)"
    cats(5) = "Organisation (System & Policy)"

    Dim menuText As String
    menuText = "SELECT A CATEGORY" & vbCrLf & vbCrLf
    menuText = menuText & "1 - People (Human Factors)" & vbCrLf
    menuText = menuText & "2 - Equipment (Resources & Equipment)" & vbCrLf
    menuText = menuText & "3 - Environment (Physical & Workplace)" & vbCrLf
    menuText = menuText & "4 - Processes (Tasks & Procedures)" & vbCrLf
    menuText = menuText & "5 - Organisation (System & Policy)" & vbCrLf
    menuText = menuText & vbCrLf & "Enter number (or 0 when finished):"

    Dim allSelected As String
    allSelected = ""

    Do
        Dim catChoice As String
        catChoice = InputBox(menuText, "PEEPO Factor Selection")

        If catChoice = "" Or catChoice = "0" Then Exit Do

        Dim catNum As Long
        catNum = Val(catChoice)

        If catNum < 1 Or catNum > 5 Then
            MsgBox "Please enter a number between 1 and 5.", vbExclamation
        Else
            Dim factorList As String
            factorList = GetFactorsForCategory(catNum)

            Dim selected As String
            selected = InputBox(factorList, cats(catNum))

            If selected <> "" Then
                Dim result As String
                result = ParseSelection(catNum, selected)
                If result <> "" Then
                    If allSelected <> "" Then
                        allSelected = allSelected & vbCrLf
                    End If
                    allSelected = allSelected & cats(catNum) & ":" & vbCrLf & result
                End If
            End If
        End If
    Loop

    If allSelected <> "" Then
        InsertFactorsIntoTable allSelected
    End If
End Sub

Private Function GetFactorsForCategory(catNum As Long) As String
    Dim s As String
    s = "Enter the numbers of applicable factors," & vbCrLf
    s = s & "separated by commas (e.g. 1,3,5):" & vbCrLf & vbCrLf

    Select Case catNum
    Case 1
        s = s & "1 - Inadequate communication between staff" & vbCrLf
        s = s & "2 - Unreasonable behaviour by worker(s)" & vbCrLf
        s = s & "3 - Unreasonable behaviour by manager(s) or supervisor(s)" & vbCrLf
        s = s & "4 - Interpersonal conflict (unresolved or escalating)" & vbCrLf
        s = s & "5 - Fatigue or burnout" & vbCrLf
        s = s & "6 - Mental stress or psychological distress" & vbCrLf
        s = s & "7 - Exposure to traumatic event or material" & vbCrLf
        s = s & "8 - Personal issues affecting work capacity" & vbCrLf
        s = s & "9 - Incivility or inappropriate behaviour tolerated by peers"
    Case 2
        s = s & "1 - Insufficient staffing levels or resources" & vbCrLf
        s = s & "2 - Inadequate or unavailable tools and resources" & vbCrLf
        s = s & "3 - Insufficient training on psychosocial hazard ID or mgmt" & vbCrLf
        s = s & "4 - Insufficient training for job demands or new tasks" & vbCrLf
        s = s & "5 - Lack of constructive feedback or performance support" & vbCrLf
        s = s & "6 - Inadequate access to employee assistance or support" & vbCrLf
        s = s & "7 - Lack of info about reporting processes or support" & vbCrLf
        s = s & "8 - Insufficient or inappropriate supervision" & vbCrLf
        s = s & "9 - Inadequate de-escalation or conflict resolution resources"
    Case 3
        s = s & "1 - Adverse environmental conditions (noise, temp, air)" & vbCrLf
        s = s & "2 - Poorly maintained or unclean facilities" & vbCrLf
        s = s & "3 - Workplace layout limiting supervision or support" & vbCrLf
        s = s & "4 - Remote or isolated work location" & vbCrLf
        s = s & "5 - Natural events restricting travel or creating uncertainty" & vbCrLf
        s = s & "6 - Limited access to communication technology" & vbCrLf
        s = s & "7 - Blurring of boundaries between work and home life" & vbCrLf
        s = s & "8 - Workplace culture tolerating unreasonable behaviour" & vbCrLf
        s = s & "9 - Lack of diversity, respect, or inclusion" & vbCrLf
        s = s & "10 - Limited opportunities for social interaction"
    Case 4
        s = s & "1 - No/inadequate policies for psychosocial hazards" & vbCrLf
        s = s & "2 - Policies or procedures not adhered to" & vbCrLf
        s = s & "3 - Procedures that lack flexibility or cannot be applied" & vbCrLf
        s = s & "4 - No/inadequate reporting and complaints processes" & vbCrLf
        s = s & "5 - Investigation/appeals not affording natural justice" & vbCrLf
        s = s & "6 - No/inadequate incident response or investigation" & vbCrLf
        s = s & "7 - Procedures that discriminate or lack fairness" & vbCrLf
        s = s & "8 - Inadequate rostering, hours, or fatigue management" & vbCrLf
        s = s & "9 - Inadequate change management procedures" & vbCrLf
        s = s & "10 - No mechanism for addressing senior mgmt behaviour"
    Case 5
        s = s & "1 - Poor leadership practices or management style" & vbCrLf
        s = s & "2 - Limited management accountability for psych hazards" & vbCrLf
        s = s & "3 - Unclear or conflicting roles, responsibilities" & vbCrLf
        s = s & "4 - Excessive work demands (cognitive, emotional, physical)" & vbCrLf
        s = s & "5 - Workplace culture" & vbCrLf
        s = s & "6 - Multiple interacting hazards not considered together" & vbCrLf
        s = s & "7 - Low job control or autonomy" & vbCrLf
        s = s & "8 - Lack of recognition or reward" & vbCrLf
        s = s & "9 - Organisational injustice (unfairness, bias)" & vbCrLf
        s = s & "10 - Inadequate consultation with workers" & vbCrLf
        s = s & "11 - Poorly managed organisational change" & vbCrLf
        s = s & "12 - Insecure work arrangements" & vbCrLf
        s = s & "13 - Inadequate monitoring and review of controls"
    End Select

    GetFactorsForCategory = s
End Function

Private Function GetFactorText(catNum As Long, factorNum As Long) As String
    Dim f As String
    f = ""
    Select Case catNum
    Case 1
        Select Case factorNum
        Case 1: f = "Inadequate communication between staff"
        Case 2: f = "Unreasonable behaviour by worker(s)"
        Case 3: f = "Unreasonable behaviour by manager(s) or supervisor(s)"
        Case 4: f = "Interpersonal conflict (unresolved or escalating)"
        Case 5: f = "Fatigue or burnout"
        Case 6: f = "Mental stress or psychological distress"
        Case 7: f = "Exposure to traumatic event or material"
        Case 8: f = "Personal issues affecting work capacity"
        Case 9: f = "Incivility or inappropriate behaviour tolerated by peers"
        End Select
    Case 2
        Select Case factorNum
        Case 1: f = "Insufficient staffing levels or resources to meet work demands"
        Case 2: f = "Inadequate or unavailable tools and resources to do the job"
        Case 3: f = "Insufficient training on psychosocial hazard identification or management"
        Case 4: f = "Insufficient training for job demands or new tasks"
        Case 5: f = "Lack of constructive feedback or performance support"
        Case 6: f = "Inadequate access to employee assistance or support services"
        Case 7: f = "Lack of information about reporting processes or support pathways"
        Case 8: f = "Insufficient or inappropriate supervision"
        Case 9: f = "Inadequate de-escalation tools or conflict resolution resources"
        End Select
    Case 3
        Select Case factorNum
        Case 1: f = "Adverse environmental conditions (noise, temperature, air quality)"
        Case 2: f = "Poorly maintained or unclean facilities"
        Case 3: f = "Workplace layout limiting supervision or support access"
        Case 4: f = "Remote or isolated work location"
        Case 5: f = "Natural events restricting travel or creating uncertainty"
        Case 6: f = "Limited access to communication technology"
        Case 7: f = "Blurring of boundaries between work and home life"
        Case 8: f = "Workplace culture tolerating unreasonable behaviour"
        Case 9: f = "Lack of diversity, respect, or inclusion in the workplace"
        Case 10: f = "Limited opportunities for social interaction during work"
        End Select
    Case 4
        Select Case factorNum
        Case 1: f = "No or inadequate policies for managing psychosocial hazards"
        Case 2: f = "Policies or procedures not adhered to"
        Case 3: f = "Procedures that cannot be applied as written or lack flexibility"
        Case 4: f = "No or inadequate reporting and complaints processes"
        Case 5: f = "Investigation or appeals process not affording natural justice"
        Case 6: f = "No or inadequate incident response or investigation procedures"
        Case 7: f = "Procedures that discriminate or lack fairness"
        Case 8: f = "Inadequate rostering, hours of work, or fatigue management procedures"
        Case 9: f = "Inadequate change management procedures"
        Case 10: f = "No mechanism for impartially addressing behaviour by senior management"
        End Select
    Case 5
        Select Case factorNum
        Case 1: f = "Poor leadership practices or management style"
        Case 2: f = "Limited management accountability for psychosocial hazards"
        Case 3: f = "Unclear or conflicting roles, responsibilities, or expectations"
        Case 4: f = "Excessive work demands (cognitive, emotional, or physical)"
        Case 5: f = "Workplace culture"
        Case 6: f = "Multiple interacting psychosocial hazards not considered together"
        Case 7: f = "Low job control or autonomy"
        Case 8: f = "Lack of recognition or reward"
        Case 9: f = "Organisational injustice (unfairness, bias, inconsistency)"
        Case 10: f = "Inadequate consultation with workers on matters affecting them"
        Case 11: f = "Poorly managed organisational change"
        Case 12: f = "Insecure work arrangements"
        Case 13: f = "Inadequate monitoring and review of existing controls"
        End Select
    End Select
    GetFactorText = f
End Function

Private Function ParseSelection(catNum As Long, selection As String) As String
    Dim parts() As String
    parts = Split(selection, ",")
    Dim result As String
    result = ""
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim num As Long
        num = Val(Trim(parts(i)))
        If num > 0 Then
            Dim factorText As String
            factorText = GetFactorText(catNum, num)
            If factorText <> "" Then
                result = result & "  - " & factorText & vbCrLf
            End If
        End If
    Next i
    ParseSelection = result
End Function

Public Sub InsertFactorsIntoTable(ByVal factorText As String)
    Dim tbl As Table
    Dim targetCell As Cell
    Dim found As Boolean
    found = False
    Dim r As Long
    Dim c As Long

    For Each tbl In ActiveDocument.Tables
        For r = 1 To tbl.Rows.Count
            For c = 1 To tbl.Columns.Count
                On Error Resume Next
                Dim cellText As String
                cellText = ""
                cellText = tbl.Cell(r, c).Range.Text
                On Error GoTo 0
                If InStr(1, cellText, "Outline any relevant work or non-work factors", vbTextCompare) > 0 Then
                    On Error Resume Next
                    Set targetCell = tbl.Cell(r + 1, 1)
                    If targetCell Is Nothing Then
                        Set targetCell = tbl.Cell(r + 1, c)
                    End If
                    On Error GoTo 0
                    If Not targetCell Is Nothing Then
                        found = True
                        GoTo FoundIt
                    End If
                End If
            Next c
        Next r
    Next tbl

FoundIt:
    If found Then
        Dim rng As Range
        Set rng = targetCell.Range
        rng.End = rng.End - 1

        Dim existingText As String
        existingText = Trim(rng.Text)

        If existingText <> "" Then
            rng.Collapse Direction:=0
            rng.InsertAfter vbCr & vbCr & factorText
        Else
            rng.Text = factorText
        End If

        MsgBox "Factors inserted successfully into the table.", vbInformation, "PEEPO"
    Else
        MsgBox "Could not find the target cell." & vbCrLf & vbCrLf & _
               "Please ensure this document contains the Nature of Concerns " & _
               "table with the row 'Outline any relevant work or non-work factors'.", _
               vbExclamation, "PEEPO"
    End If
End Sub
"""

THIS_DOC_CODE = ('Attribute VB_Name = "ThisDocument"\r\n'
                 'Attribute VB_Base = "1Normal.ThisDocument"\r\n'
                 'Attribute VB_GlobalNameSpace = False\r\n'
                 'Attribute VB_Creatable = False\r\n'
                 'Attribute VB_PredeclaredId = True\r\n'
                 'Attribute VB_Exposed = True\r\n'
                 'Attribute VB_TemplateDerived = True\r\n'
                 'Attribute VB_Customizable = True\r\n')


def main():
    print("Building vbaProject.bin...")

    dir_stream = build_dir_stream(MODULE_NAME)
    this_doc_stream = make_vba_module_stream(THIS_DOC_CODE)
    module_stream = make_vba_module_stream(VBA_CODE)

    vba_project_stream = bytes([
        0xCC, 0x61, 0x00, 0x00, 0x00, 0x01, 0x00
    ])

    project_text = (
        'ID="{00000000-0000-0000-0000-000000000001}"\r\n'
        'Document=ThisDocument/&H00000000\r\n'
        f'Module={MODULE_NAME}\r\n'
        'Name="VBAProject"\r\n'
        'HelpContextID="0"\r\n'
        'VersionCompatible32="393222000"\r\n'
        'CMG="0000"\r\n'
        'DPB="0000"\r\n'
        'GC="0000"\r\n'
        '\r\n'
        '[Host Extender Info]\r\n'
        '&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000\r\n'
    ).encode('ascii')

    projectwm = bytearray()
    for name in ["ThisDocument", MODULE_NAME]:
        projectwm.extend(name.encode('ascii') + b'\x00')
        projectwm.extend(name.encode('utf-16-le') + b'\x00\x00')
    projectwm.extend(b'\x00')

    vba_bin = write_cfb({
        'VBA/dir': dir_stream,
        'VBA/_VBA_PROJECT': vba_project_stream,
        'VBA/ThisDocument': this_doc_stream,
        f'VBA/{MODULE_NAME}': module_stream,
        'PROJECT': project_text,
        'PROJECTwm': bytes(projectwm),
    })

    print(f"vbaProject.bin size: {len(vba_bin)} bytes")

    print("Injecting VBA into document...")
    inject_vba_into_docx(SOURCE_DOCX, OUTPUT_DOCM, vba_bin)

    print(f"Created: {OUTPUT_DOCM}")

    # Verify
    import olefile
    with zipfile.ZipFile(OUTPUT_DOCM, 'r') as z:
        vba_data = z.read('word/vbaProject.bin')
        ole = olefile.OleFileIO(io.BytesIO(vba_data))
        print("\nOLE streams:")
        for s in ole.listdir():
            path = '/'.join(s)
            print(f"  {path} ({ole.get_size(path)} bytes)")
        ole.close()

    print("\nDone! Open in Word, press Alt+F8, run 'SelectPEEPOFactors'")


if __name__ == "__main__":
    main()
