import xml.etree.ElementTree as ET
import re
import copy

class XMLAnalyzer:
    def __init__(self):
        self.raw_xml_string = ""
        self.safe_header = None
        self.tree = None
        self.root = None
        self.history = []

    def save_state(self):
        if self.root is not None:
            self.history.append(ET.tostring(self.root, encoding="UTF-8").decode("UTF-8"))
            if len(self.history) > 15:
                self.history.pop(0)

    def undo(self):
        if self.history:
            old_xml = self.history.pop()
            self.root = ET.fromstring(old_xml)
            return True
        return False

    def load_from_string(self, xml_string):
        safe_xml_string = re.sub(r'&(?!amp;|lt;|gt;|quot;|apos;|#)', '&amp;', xml_string)
        self.raw_xml_string = safe_xml_string
        self.safe_header = self._extract_header_block_from_string(safe_xml_string)
        self.root = ET.fromstring(safe_xml_string)
        self.history.clear()
        print("File loaded into RAM and sanitized successfully.")

    def _extract_header_block_from_string(self, xml_string):
        pattern = r"(?s)(<form.*?</pipPrefs>)"
        match = re.search(pattern, xml_string)
        if match: return match.group(1)
        return ""

    def get_final_xml_string(self):
        modified_xml_bytes = ET.tostring(self.root, encoding="UTF-8", xml_declaration=True)
        modified_xml_string = modified_xml_bytes.decode("UTF-8")
        if self.safe_header:
            pattern = r"(?s)(<form.*?(?:</pipPrefs>|<pipPrefs\s*/>))"
            if re.search(pattern, modified_xml_string):
                return re.sub(pattern, lambda m: self.safe_header, modified_xml_string, count=1)
        return modified_xml_string

    def ensure_accent_row(self):
        if self.root is None: return
        rows_node = self.root.find(".//query/rows")
        if rows_node is None: return
        segments = rows_node.findall("segment")
        if not segments: return
        
        seg0 = segments[0]
        is_seg0_formula = seg0.find(".//formula") is not None
        is_seg0_spacer = (seg0.get("size", "") == "-4")
        
        def create_spacer(base_seg):
            new_seg = copy.deepcopy(base_seg)
            new_seg.set("size", "-4")
            for dim in new_seg.findall("dimension"):
                for child in list(dim): dim.remove(child) 
                ET.SubElement(dim, "formula", ordinal="1.0", dataType="0", label="Formula Label", formulaValue="0")
            return new_seg

        if is_seg0_formula and is_seg0_spacer:
            formula_tag = seg0.find(".//formula")
            if formula_tag is not None and not formula_tag.get("formulaValue"):
                formula_tag.set("formulaValue", "0")
        elif is_seg0_formula and not is_seg0_spacer:
            needs_spacer = True
            if len(segments) > 1:
                seg1 = segments[1]
                if seg1.find(".//formula") is not None and seg1.get("size", "") == "-4":
                    needs_spacer = False
            if needs_spacer: rows_node.insert(1, create_spacer(seg0))
        else:
            rows_node.insert(0, create_spacer(seg0))

    def get_rowcols(self):
        if self.root is None: return {"rows": [], "columns": []}
            
        def parse_segments(xml_path):
            container_list = []
            parent_node = self.root.find(xml_path)
            if parent_node is None: return []
                
            for s_idx, segment in enumerate(parent_node.findall("segment")):
                dim_size = segment.get("size", segment.get("width", segment.get("height", "")))
                combos = [{}] 
                for dimension in segment.findall("dimension"):
                    dim_name = dimension.get("name", "")
                    dim_items = []
                    for child in dimension:
                        if child.tag == "formula":
                            dim_items.append({"name": child.get("label", "Formula"), "type": "FORMULA"})
                        elif child.tag == "function":
                            if child.get("exclude") == "true": continue
                            mbr = child.find("member")
                            raw_name = mbr.get("name") if mbr is not None else ""
                            dim_items.append({"name": raw_name, "_debug_name": f"{child.get('name')}({raw_name})", "type": "FUNCTION"})
                        elif child.tag == "member":
                            if child.get("exclude") == "true": continue
                            dim_items.append({"name": child.get("name", ""), "type": "MEMBER"})
                            
                    if dim_items:
                        new_combos = []
                        for combo in combos:
                            for item in dim_items:
                                new_c = combo.copy()
                                new_c[dim_name] = item["name"] 
                                new_c["_display_name"] = item["name"] 
                                new_c["_type"] = item["type"]
                                new_c["_size"] = dim_size
                                new_combos.append(new_c)
                        combos = new_combos
                for c in combos:
                    if c: 
                        c["_segment_idx"] = s_idx + 1 
                        container_list.append(c)
            return container_list
        return {"rows": parse_segments(".//query/rows"), "columns": parse_segments(".//query/columns")}

    # UPGRADED: Extracts txtColor dynamically
    def get_format_map(self):
        format_map = {}
        if self.root is None: return format_map

        colors_dict = {c.get("id"): f"#{int(c.get('R', '0')):02X}{int(c.get('G', '0')):02X}{int(c.get('B', '0')):02X}"
                       for c in self.root.findall(".//values/colors/color")}
        style_to_hex = {}
        
        for style in self.root.findall(".//cellStyles/cellStyle"):
            bg = style.find(".//backColor")
            fg = style.find(".//txtColor")
            
            bg_hex = colors_dict.get(bg.get("id")) if bg is not None else None
            fg_hex = colors_dict.get(fg.get("id")) if fg is not None else None
            
            style_to_hex[style.get("id")] = {"bg": bg_hex, "fg": fg_hex}

        for dvr in self.root.findall(".//dataValidationRules/dataValidationRule"):
            try:
                r_loc = int(float(dvr.get("rowLocation", "0")))
                c_loc = int(float(dvr.get("colLocation", "0")))
                cond = dvr.find("dataValidationCond")
                
                if cond is not None and cond.get("styleId") in style_to_hex:
                    format_map[(r_loc, c_loc)] = style_to_hex[cond.get("styleId")]
            except ValueError: 
                pass

        return format_map

    def strip_legacy_formatting(self):
        if self.root is None: return False
        self.save_state() 
        
        tuples_bucket = self.root.find(".//formFormattings/formFormatting/dataCellMbrTuples")
        if tuples_bucket is not None: tuples_bucket.clear()
            
        dvr_bucket = self.root.find(".//dataValidationRules")
        if dvr_bucket is not None:
            dvrs_to_remove = []
            for dvr in dvr_bucket.findall("dataValidationRule"):
                if dvr.get("customStyle") == "true":
                    cond = dvr.find("dataValidationCond")
                    if cond is not None and cond.get("type") == "8" and (cond.find("compareValue") is None or cond.find("compareValue").get("value") == ""):
                        dvrs_to_remove.append(dvr)
                    else:
                        dvr.set("customStyle", "false")
                        if cond is not None:
                            cond.attrib.pop("styleId", None)
                            cond.attrib.pop("bgColor", None)
            for d in dvrs_to_remove: dvr_bucket.remove(d)
        return True

    def apply_master_formatting(self):
        if self.root is None: return False
        self.save_state() 
        
        temp_history = self.history.copy()
        self.strip_legacy_formatting()
        self.history = temp_history 
        
        self.ensure_accent_row()
        grid_data = self.get_rowcols()
        
        self.setup_formatting_foundation()
        self.ensure_txt_formats()
        
        dark_blue_id = self.add_new_color("11", "37", "49", ignore_history=True)
        light_blue_id = self.add_new_color("240", "248", "255", ignore_history=True)
        white_id = self.add_new_color("255", "255", "255", ignore_history=True)
        orange_id = self.add_new_color("255", "140", "0", ignore_history=True) 
        
        border_ids = self.inject_standard_borders()
        col_style_id = self.add_advanced_cell_style(bg_color_id=dark_blue_id, txt_color_id=white_id, is_bold=True, border_ids=border_ids)
        row_style_id = self.add_advanced_cell_style(bg_color_id=light_blue_id)
        orange_style_id = self.add_advanced_cell_style(bg_color_id=orange_id) 

        max_col_seg = max((c.get("_segment_idx", 1) for c in grid_data["columns"]), default=0)
        max_row_seg = max((r.get("_segment_idx", 1) for r in grid_data["rows"]), default=0)

        for c_idx in range(max_col_seg):
            self.add_location_dvr(row_loc=0.0, col_loc=c_idx+1, style_id=col_style_id, hex_color="0B2531", rule_name=f"Column {c_idx+1} Header Format")

        for r_idx in range(1, max_row_seg + 1):
            row_info = next((r for r in grid_data["rows"] if r.get("_segment_idx") == r_idx), None)
            if not row_info: continue
            
            is_formula = (row_info.get("_type") == "FORMULA")
            is_spacer = (row_info.get("_size") == "-4")
            
            if r_idx == 1 and is_formula and not is_spacer:
                self.add_location_dvr(row_loc=r_idx, col_loc=0.0, style_id=col_style_id, hex_color="0B2531", rule_name=f"Row {r_idx} Header Format (Formula)")
                self.add_location_dvr(row_loc=r_idx, col_loc=-1.0, style_id=col_style_id, hex_color="0B2531", rule_name=f"Row {r_idx} Data Format (Formula)")
            elif is_formula and is_spacer:
                self.add_location_dvr(row_loc=r_idx, col_loc=0.0, style_id=orange_style_id, hex_color="FF8C00", rule_name=f"Row {r_idx} Header Format (Accent Line)")
                self.add_location_dvr(row_loc=r_idx, col_loc=-1.0, style_id=orange_style_id, hex_color="FF8C00", rule_name=f"Row {r_idx} Data Format (Accent Line)")
            else:
                self.add_location_dvr(row_loc=r_idx, col_loc=0.0, style_id=row_style_id, hex_color="F0F8FF", rule_name=f"Row {r_idx} Header Format (Data)")

        return True

    def get_detailed_colors(self):
        if self.root is None: return []

        color_data = {}
        colors_bucket = self.root.find(".//values/colors")
        if colors_bucket is not None:
            for c in colors_bucket.findall("color"):
                c_id = c.get("id")
                hex_val = self.rgb_to_hex([c.get("R","0"), c.get("G","0"), c.get("B","0")])
                color_data[c_id] = {"hex": hex_val, "styles": [], "locations": []}

        style_to_color = {}
        styles_bucket = self.root.find(".//cellStyles")
        if styles_bucket is not None:
            for s in styles_bucket.findall("cellStyle"):
                bg = s.find(".//backColor")
                if bg is not None and bg.get("id") in color_data:
                    c_id = bg.get("id")
                    s_id = s.get("id")
                    style_to_color[s_id] = c_id
                    color_data[c_id]["styles"].append(s_id)

        dvr_bucket = self.root.find(".//dataValidationRules")
        if dvr_bucket is not None:
            for dvr in dvr_bucket.findall("dataValidationRule"):
                cond = dvr.find("dataValidationCond")
                if cond is not None:
                    s_id = cond.get("styleId")
                    if s_id in style_to_color:
                        c_id = style_to_color[s_id]
                        r_loc = dvr.get("rowLocation", "0")
                        c_loc = dvr.get("colLocation", "0")
                        loc_dict = {"type": "dvr", "r": r_loc, "c": c_loc, "display": f"R:{int(float(r_loc))}, C:{int(float(c_loc))}"}
                        color_data[c_id]["locations"].append(loc_dict)

        tuples_bucket = self.root.find(".//formFormattings/formFormatting/dataCellMbrTuples")
        if tuples_bucket is not None:
            for t in tuples_bucket.findall("dataCellMbrTuple"):
                c_id_node = t.find("cellStyleId")
                if c_id_node is not None and c_id_node.text in style_to_color:
                    c_id = style_to_color[c_id_node.text]
                    mbr = t.find(".//mbr")
                    if mbr is not None:
                        loc_dict = {"type": "tuple", "mbr": mbr.get("name"), "display": f"Mbr: {mbr.get('name')}"}
                        color_data[c_id]["locations"].append(loc_dict)

        results = []
        for c_id, info in color_data.items():
            unique_locs = []
            seen = set()
            for loc in info["locations"]:
                if loc["display"] not in seen:
                    seen.add(loc["display"])
                    unique_locs.append(loc)
            results.append({"id": c_id, "hex": info["hex"], "locations_data": unique_locs})
            
        return results

    def remove_specific_color_link(self, color_id, loc_data):
        if self.root is None: return
        self.save_state() 
        
        styles_bucket = self.root.find(".//cellStyles")
        style_ids = []
        if styles_bucket is not None:
            for style in styles_bucket.findall("cellStyle"):
                bg = style.find(".//backColor")
                if bg is not None and bg.get("id") == str(color_id):
                    style_ids.append(style.get("id"))
                    
        if loc_data["type"] == "dvr":
            dvr_bucket = self.root.find(".//dataValidationRules")
            if dvr_bucket is not None:
                for dvr in dvr_bucket.findall("dataValidationRule"):
                    if dvr.get("rowLocation") == loc_data["r"] and dvr.get("colLocation") == loc_data["c"]:
                        cond = dvr.find("dataValidationCond")
                        if cond is not None and cond.get("styleId") in style_ids:
                            dvr_bucket.remove(dvr)
                            
        elif loc_data["type"] == "tuple":
            tuples_bucket = self.root.find(".//formFormattings/formFormatting/dataCellMbrTuples")
            if tuples_bucket is not None:
                for t in tuples_bucket.findall("dataCellMbrTuple"):
                    c_id_node = t.find("cellStyleId")
                    if c_id_node is not None and c_id_node.text in style_ids:
                        mbr = t.find(".//mbr")
                        if mbr is not None and mbr.get("name") == loc_data["mbr"]:
                            tuples_bucket.remove(t)

    def remove_color_and_usages(self, color_id):
        if self.root is None: return
        self.save_state() 
        
        styles_to_remove = set()
        styles_bucket = self.root.find(".//cellStyles")
        if styles_bucket is not None:
            for style in styles_bucket.findall("cellStyle"):
                bg = style.find(".//backColor")
                if bg is not None and bg.get("id") == str(color_id):
                    styles_to_remove.add(style.get("id"))
                    styles_bucket.remove(style)
        dvr_bucket = self.root.find(".//dataValidationRules")
        if dvr_bucket is not None:
            for dvr in dvr_bucket.findall("dataValidationRule"):
                cond = dvr.find("dataValidationCond")
                if cond is not None and cond.get("styleId") in styles_to_remove:
                    dvr_bucket.remove(dvr)
        tuples_bucket = self.root.find(".//formFormattings/formFormatting/dataCellMbrTuples")
        if tuples_bucket is not None:
            for t in tuples_bucket.findall("dataCellMbrTuple"):
                c_id_node = t.find("cellStyleId")
                if c_id_node is not None and c_id_node.text in styles_to_remove:
                    tuples_bucket.remove(t)
        colors_bucket = self.root.find(".//values/colors")
        if colors_bucket is not None:
            for c in colors_bucket.findall("color"):
                if c.get("id") == str(color_id):
                    colors_bucket.remove(c)

    def inject_colors(self, color_list):
        if self.root == None: return
        self.save_state() 
        
        color_map = {str(id_data): color_data for id_data, color_data in color_list}
        colors = self.root.find(".//values/colors")  
        if colors is not None:
            for color in colors.findall("color"):
                xml_id = str(color.get("id"))
                if xml_id in color_map:
                    hex_value = color_map[xml_id]
                    rgb_list = self.hex_to_rgb([hex_value])
                    color.set("R", str(rgb_list[0][0]))
                    color.set("G", str(rgb_list[0][1]))
                    color.set("B", str(rgb_list[0][2]))

    def setup_formatting_foundation(self):
        if self.root is None: return False
        ff_parent = self.root.find(".//formFormattings")
        if ff_parent is None: ff_parent = ET.SubElement(self.root, "formFormattings")
        ff = ff_parent.find("formFormatting")
        if ff is None: ff = ET.SubElement(ff_parent, "formFormatting", designTime="true", userName="[CURRENT_USER]", displayOptions="-2147483646")
        for bucket in ["dataCellMbrTuples", "cellStyles", "columnRowSizes"]:
            if ff.find(bucket) is None: ET.SubElement(ff, bucket)
        values = ff.find("values")
        if values is None:
            values = ET.SubElement(ff, "values")
            txt_frmts = ET.SubElement(values, "txtFrmts")
            ET.SubElement(txt_frmts, "txtFrmt", id="1").text = "Bold"
            ET.SubElement(txt_frmts, "txtFrmt", id="2").text = "Underline"
            ET.SubElement(txt_frmts, "txtFrmt", id="3").text = "StrikeThrough"
            ET.SubElement(values, "colors")
        objs = ff.find("objs")
        if objs is None:
            objs = ET.SubElement(ff, "objs")
            ET.SubElement(objs, "numFrmts")
            ET.SubElement(objs, "borders")
        return False

    def ensure_txt_formats(self):
        txt_bucket = self.root.find(".//formFormattings/formFormatting/values/txtFrmts")
        if txt_bucket is not None:
            formats = {"0": "Italic", "1": "Bold", "2": "Underline", "3": "StrikeThrough"}
            for f_id, f_text in formats.items():
                if txt_bucket.find(f"txtFrmt[@id='{f_id}']") is None:
                    ET.SubElement(txt_bucket, "txtFrmt", id=f_id).text = f_text

    def get_next_available_id(self):
        highest_id = 32767
        if self.root is not None:
            ff_node = self.root.find(".//formFormattings/formFormatting")
            if ff_node is not None:
                for elem in ff_node.findall(".//*[@id]"):
                    try: highest_id = max(highest_id, int(elem.get("id")))
                    except ValueError: pass 
                for elem in ff_node.findall(".//id"):
                    try:
                        if elem.text: highest_id = max(highest_id, int(elem.text.strip()))
                    except ValueError: pass
        return highest_id + 1

    def add_new_color(self, r, g, b, ignore_history=False):
        if not ignore_history: self.save_state()
            
        colors_bucket = self.root.find(".//formFormattings/formFormatting/values/colors")
        if colors_bucket is not None:
            new_id = self.get_next_available_id()
            ET.SubElement(colors_bucket, "color", id=str(new_id), R=str(r), G=str(g), B=str(b))
            return new_id
        return None

    def inject_standard_borders(self):
        hardcoded_xml = """<borders>
            <border><id>32768</id><color R="255" G="255" B="255" /><placement>Top</placement><style>solid</style><width>0.4</width></border>
            <border><id>32769</id><color R="255" G="255" B="255" /><placement>Right</placement><style>solid</style><width>0.4</width></border>
            <border><id>32770</id><color R="255" G="255" B="255" /><placement>Bottom</placement><style>solid</style><width>0.4</width></border>
            <border><id>32771</id><color R="255" G="255" B="255" /><placement>Left</placement><style>solid</style><width>0.4</width></border>
        </borders>"""
        new_borders = ET.fromstring(hardcoded_xml)
        for elem in self.root.iter():
            if 'borders' in elem.tag.lower():
                elem.clear() 
                for new_border in new_borders: elem.append(new_border)
                return [32768, 32769, 32770, 32771]
        return []

    def add_advanced_cell_style(self, bg_color_id, txt_color_id=None, is_bold=False, border_ids=None):
        style_id = self.get_next_available_id()
        styles_bucket = self.root.find(".//formFormattings/formFormatting/cellStyles")
        if styles_bucket is not None:
            cell_style = ET.SubElement(styles_bucket, "cellStyle", id=str(style_id))
            objs_node = ET.SubElement(cell_style, "objs")
            if border_ids:
                for b_id in border_ids: ET.SubElement(objs_node, "obj", type="border", id=str(b_id))
            values_node = ET.SubElement(cell_style, "cellStyleValues")
            ET.SubElement(values_node, "font", id="32768") 
            ET.SubElement(values_node, "readOnly").text = "false"
            ET.SubElement(values_node, "backColor", id=str(bg_color_id))
            if txt_color_id: ET.SubElement(values_node, "txtColor", id=str(txt_color_id))
            ET.SubElement(values_node, "wordWrap").text = "false"
            if is_bold: ET.SubElement(values_node, "format", id="1")
            return style_id
        return None

    def add_location_dvr(self, row_loc, col_loc, style_id, hex_color, rule_name="Auto Format Rule"):
        clean_hex = hex_color.replace("#", "")
        decimal_color = str(int(clean_hex, 16))
        dvr_bucket = self.root.find(".//dataValidationRules")
        if dvr_bucket is None: return
        rule = ET.SubElement(dvr_bucket, "dataValidationRule", position="1", name=rule_name, description="Automated layout styling rule generated by the EPM XML Multi-Tool. © Adetth Raju", enabled="true", customStyle="true", rowLocation=str(float(row_loc)), colLocation=str(float(col_loc)))
        cond = ET.SubElement(rule, "dataValidationCond", toolTip="", groupOpenNestingLevel="0", operator="0", displayMessageInDVPane="false", honorPmRules="false", negate="false", groupCloseNestingLevel="0", position="1", styleId=str(style_id), type="8", bgColor=decimal_color, Valid="true", logicalOperator="0")
        ET.SubElement(cond, "compareValue", type="6", value="")
        ET.SubElement(cond, "compareToValue", type="0", value="")

    def inject_cell_color(self, r_loc, c_loc, hex_color):
        if self.root is None: return False
        self.save_state() 

        clean_hex = hex_color.replace("#", "").upper()
        
        colors_bucket = self.root.find(".//values/colors")
        color_id = None
        if colors_bucket is not None:
            for c in colors_bucket.findall("color"):
                if self.rgb_to_hex([c.get("R","0"), c.get("G","0"), c.get("B","0")]) == clean_hex:
                    color_id = c.get("id")
                    break
        
        if not color_id:
            rgb = self.hex_to_rgb([clean_hex])[0]
            color_id = self.add_new_color(rgb[0], rgb[1], rgb[2], ignore_history=True)

        styles_bucket = self.root.find(".//cellStyles")
        style_id = None
        if styles_bucket is not None:
            for s in styles_bucket.findall("cellStyle"):
                bg = s.find(".//backColor")
                if bg is not None and bg.get("id") == str(color_id):
                    style_id = s.get("id")
                    break
                    
        if not style_id:
            style_id = self.add_advanced_cell_style(bg_color_id=color_id)

        dvr_bucket = self.root.find(".//dataValidationRules")
        if dvr_bucket is None: return False
            
        decimal_color = str(int(clean_hex, 16))
        rule_found = False
        
        for dvr in dvr_bucket.findall("dataValidationRule"):
            if dvr.get("rowLocation") == str(float(r_loc)) and dvr.get("colLocation") == str(float(c_loc)):
                dvr.set("customStyle", "true")
                cond = dvr.find("dataValidationCond")
                if cond is not None:
                    cond.set("styleId", str(style_id))
                    cond.set("bgColor", decimal_color)
                    cond.set("type", "8") 
                rule_found = True
                break
                
        if not rule_found:
            self.add_location_dvr(row_loc=r_loc, col_loc=c_loc, style_id=style_id, hex_color=clean_hex, rule_name=f"Interactive Cell [{r_loc}, {c_loc}]")
            
        return True

    @staticmethod
    def rgb_to_hex(color_list):
        return f"{int(color_list[0]):02X}{int(color_list[1]):02X}{int(color_list[2]):02X}".upper()

    @staticmethod
    def hex_to_rgb(color_list):
        return_color_list = []
        for hex_str in color_list:
            clean_hex = hex_str.replace("#", "")
            return_color_list.append([str(int(clean_hex[0:2], 16)), str(int(clean_hex[2:4], 16)), str(int(clean_hex[4:6], 16))])
        return return_color_list