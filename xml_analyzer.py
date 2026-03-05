import xml.etree.ElementTree as ET
import re
import inspect
import copy

class XMLAnalyzer:
    def __init__(self):
        func_name = inspect.currentframe().f_code.co_name
        print(f"I am currently running inside: {func_name}")
        self.raw_xml_string = ""
        self.safe_header = None
        self.tree = None
        self.root = None

    # ==========================================
    # --- WEB-FRIENDLY RAM IO FUNCTIONS ---
    # ==========================================
    def load_from_string(self, xml_string):
        func_name = inspect.currentframe().f_code.co_name
        print(f"I am currently running inside: {func_name}")
        
        self.raw_xml_string = xml_string
        self.safe_header = self._extract_header_block_from_string(xml_string)
        self.root = ET.fromstring(xml_string)
        print("File loaded into RAM successfully.")

    def _extract_header_block_from_string(self, xml_string):
        pattern = r"(?s)(<form.*?</pipPrefs>)"
        match = re.search(pattern, xml_string)
        if match:
            print("Header block successfully captured to RAM.")
            return match.group(1)
        return ""

    def get_final_xml_string(self):
        modified_xml_bytes = ET.tostring(self.root, encoding="UTF-8", xml_declaration=True)
        modified_xml_string = modified_xml_bytes.decode("UTF-8")
        
        if self.safe_header:
            pattern = r"(?s)(<form.*?(?:</pipPrefs>|<pipPrefs\s*/>))"
            if re.search(pattern, modified_xml_string):
                final_content = re.sub(pattern, lambda m: self.safe_header, modified_xml_string, count=1)
                return final_content
        return modified_xml_string
    # ==========================================

    # --- NEW: STRUCTURAL INJECTION ENGINE ---
    def ensure_accent_row(self):
        func_name = inspect.currentframe().f_code.co_name
        print(f"I am currently running inside: {func_name}")
        
        if self.root is None: return
        
        rows_node = self.root.find(".//query/rows")
        if rows_node is None: return
        
        segments = rows_node.findall("segment")
        if not segments: return
        
        seg0 = segments[0]
        is_seg0_formula = seg0.find(".//formula") is not None
        size_seg0 = seg0.get("size", "")
        is_seg0_spacer = (size_seg0 == "-4")
        
        def create_spacer(base_seg):
            new_seg = copy.deepcopy(base_seg)
            new_seg.set("size", "-4")
            for dim in new_seg.findall("dimension"):
                for child in list(dim): dim.remove(child) 
                # THE FIX: Add "0" to formulaValue so EPM cannot suppress it!
                ET.SubElement(dim, "formula", ordinal="1.0", dataType="0", label="Formula Label", formulaValue="0")
            return new_seg

        if is_seg0_formula and is_seg0_spacer:
            # Rule 1: Ensure existing spacers have the anti-suppression "0"
            formula_tag = seg0.find(".//formula")
            if formula_tag is not None and not formula_tag.get("formulaValue"):
                formula_tag.set("formulaValue", "0")
            print("Accent row exists. Verified anti-suppression.")
            
        elif is_seg0_formula and not is_seg0_spacer:
            needs_spacer = True
            if len(segments) > 1:
                seg1 = segments[1]
                if seg1.find(".//formula") is not None and seg1.get("size", "") == "-4":
                    needs_spacer = False
            
            if needs_spacer:
                new_spacer = create_spacer(seg0)
                rows_node.insert(1, new_spacer)
                print("Injected accent row at index 1.")
                
        else:
            new_spacer = create_spacer(seg0)
            rows_node.insert(0, new_spacer)
            print("Injected accent row at index 0.")

    def get_rowcols(self):
        func_name = inspect.currentframe().f_code.co_name
        print(f"I am currently running inside: {func_name}")
        if self.root is None:
            return {"rows": [], "columns": []}
            
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
                            # Safety check: if a function is ever excluded
                            if child.get("exclude") == "true": continue
                            
                            mbr = child.find("member")
                            raw_name = mbr.get("name") if mbr is not None else ""
                            full_func = f"{child.get('name')}({raw_name})"
                            dim_items.append({"name": raw_name, "_debug_name": full_func, "type": "FUNCTION"})
                        elif child.tag == "member":
                            # THE FIX: Completely ignore any excluded members so the grid maps cleanly
                            if child.get("exclude") == "true":
                                continue
                                
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

        return {
            "rows": parse_segments(".//query/rows"),
            "columns": parse_segments(".//query/columns")
        }

    def get_format_map(self):
        func_name = inspect.currentframe().f_code.co_name
        print(f"I am currently running inside: {func_name}")
        
        format_map = {}
        if self.root is None: return format_map

        colors_dict = {c.get("id"): f"#{int(c.get('R', '0')):02X}{int(c.get('G', '0')):02X}{int(c.get('B', '0')):02X}"
                       for c in self.root.findall(".//values/colors/color")}
        style_to_hex = {}
        
        for style in self.root.findall(".//cellStyles/cellStyle"):
            bg = style.find(".//backColor")
            if bg is not None and bg.get("id") in colors_dict:
                style_to_hex[style.get("id")] = colors_dict[bg.get("id")]

        for dvr in self.root.findall(".//dataValidationRules/dataValidationRule"):
            try:
                r_loc = int(float(dvr.get("rowLocation", "0")))
                c_loc = int(float(dvr.get("colLocation", "0")))
                cond = dvr.find("dataValidationCond")
                
                if cond is not None and cond.get("styleId") in style_to_hex:
                    hex_color = style_to_hex[cond.get("styleId")]
                    format_map[(r_loc, c_loc)] = hex_color
            except ValueError: 
                pass

        return format_map

    def apply_master_formatting(self):
        func_name = inspect.currentframe().f_code.co_name
        print(f"I am currently running inside(2): {func_name}")
        if self.root is None: return False
        
        # 1. Mutate the structural XML tree FIRST
        self.ensure_accent_row()
        
        # 2. Extract the fresh layout SECOND
        grid_data = self.get_rowcols()
        
        self.setup_formatting_foundation()
        self.ensure_txt_formats()
        
        dark_blue_id = self.add_new_color("11", "37", "49")
        light_blue_id = self.add_new_color("240", "248", "255")
        white_id = self.add_new_color("255", "255", "255")
        orange_id = self.add_new_color("255", "140", "0") # The exact EPM Orange
        
        border_ids = self.inject_standard_borders()
        
        col_style_id = self.add_advanced_cell_style(bg_color_id=dark_blue_id, txt_color_id=white_id, is_bold=True, border_ids=border_ids)
        row_style_id = self.add_advanced_cell_style(bg_color_id=light_blue_id)
        orange_style_id = self.add_advanced_cell_style(bg_color_id=orange_id) 

        dvr_bucket = self.root.find(".//dataValidationRules")
        if dvr_bucket is not None:
            dvrs_to_remove = [d for d in dvr_bucket.findall("dataValidationRule") if d.get("name") == "Auto Format Rule"]
            for d in dvrs_to_remove: dvr_bucket.remove(d)

        tuples_bucket = self.root.find(".//formFormattings/formFormatting/dataCellMbrTuples")
        if tuples_bucket is not None:
            tuples_bucket.clear() 

        max_col_seg = max((c.get("_segment_idx", 1) for c in grid_data["columns"]), default=0)
        max_row_seg = max((r.get("_segment_idx", 1) for r in grid_data["rows"]), default=0)

        # 1. Paint ALL Column Metadata Dark Blue
        for c_idx in range(max_col_seg):
            self.add_location_dvr(row_loc=0.0, col_loc=c_idx+1, style_id=col_style_id, hex_color="0B2531")

        # 2. Paint Row Metadata dynamically based on rules
        for r_idx in range(1, max_row_seg + 1):
            
            # Find the properties of this specific structural segment
            row_info = next((r for r in grid_data["rows"] if r.get("_segment_idx") == r_idx), None)
            if not row_info: continue
            
            is_formula = (row_info.get("_type") == "FORMULA")
            is_spacer = (row_info.get("_size") == "-4")
            
            if r_idx == 1 and is_formula and not is_spacer:
                # Rule 2 formatting: Apply Bold/DarkBlue to both header and data
                self.add_location_dvr(row_loc=r_idx, col_loc=0.0, style_id=col_style_id, hex_color="0B2531")
                self.add_location_dvr(row_loc=r_idx, col_loc=-1.0, style_id=col_style_id, hex_color="0B2531")
                
            elif is_formula and is_spacer:
                # Accent Line formatting: Apply Orange to header and data
                self.add_location_dvr(row_loc=r_idx, col_loc=0.0, style_id=orange_style_id, hex_color="FF8C00")
                self.add_location_dvr(row_loc=r_idx, col_loc=-1.0, style_id=orange_style_id, hex_color="FF8C00")
                
            else:
                # Standard formatting: Light blue for header, blank for data
                self.add_location_dvr(row_loc=r_idx, col_loc=0.0, style_id=row_style_id, hex_color="F0F8FF")

        print("Master DVR formatting complete!")
        return True

    def get_colors(self):
        func_name = inspect.currentframe().f_code.co_name
        print(f"I am currently running inside: {func_name}")
        if self.root is None:
            return []
            
        color_data = []
        colors = self.root.find(".//values/colors")
        
        if colors is not None:
            for color in colors.findall("color"):
                color_id = color.get("id")
                r, g, b = color.get("R", "0"), color.get("G", "0"), color.get("B", "0")
                hex_val = self.rgb_to_hex([r, g, b])
                color_data.append((color_id, hex_val))
        return color_data

    def inject_colors(self, color_list):
        func_name = inspect.currentframe().f_code.co_name
        print(f"I am currently running inside: {func_name}")
        if self.root == None: return

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
        if ff_parent is None:
            ff_parent = ET.SubElement(self.root, "formFormattings")

        ff = ff_parent.find("formFormatting")
        if ff is None:
            ff = ET.SubElement(ff_parent, "formFormatting", designTime="true", userName="[CURRENT_USER]", displayOptions="-2147483646")

        tuples = ff.find("dataCellMbrTuples")
        if tuples is not None and len(list(tuples)) > 0:
            return True 

        for bucket in ["dataCellMbrTuples", "cellStyles", "columnRowSizes"]:
            if ff.find(bucket) is None:
                ET.SubElement(ff, bucket)

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
                    txt = ET.SubElement(txt_bucket, "txtFrmt", id=f_id)
                    txt.text = f_text

    def get_next_available_id(self):
        highest_id = 32767
        if self.root is not None:
            ff_node = self.root.find(".//formFormattings/formFormatting")
            if ff_node is not None:
                for elem in ff_node.findall(".//*[@id]"):
                    try:
                        highest_id = max(highest_id, int(elem.get("id")))
                    except ValueError: pass 
                for elem in ff_node.findall(".//id"):
                    try:
                        if elem.text:
                            highest_id = max(highest_id, int(elem.text.strip()))
                    except ValueError: pass
        return highest_id + 1

    def add_new_color(self, r, g, b):
        new_id = self.get_next_available_id()
        colors_bucket = self.root.find(".//formFormattings/formFormatting/values/colors")
        if colors_bucket is not None:
            new_color = ET.SubElement(colors_bucket, "color", id=str(new_id), R=str(r), G=str(g), B=str(b))
            return new_id
        return None

    def inject_standard_borders(self):
        hardcoded_xml = """
        <borders>
            <border><id>32768</id><color R="255" G="255" B="255" /><placement>Top</placement><style>solid</style><width>0.4</width></border>
            <border><id>32769</id><color R="255" G="255" B="255" /><placement>Right</placement><style>solid</style><width>0.4</width></border>
            <border><id>32770</id><color R="255" G="255" B="255" /><placement>Bottom</placement><style>solid</style><width>0.4</width></border>
            <border><id>32771</id><color R="255" G="255" B="255" /><placement>Left</placement><style>solid</style><width>0.4</width></border>
        </borders>
        """
        new_borders = ET.fromstring(hardcoded_xml)
        for elem in self.root.iter():
            if 'borders' in elem.tag.lower():
                elem.clear() 
                for new_border in new_borders:
                    elem.append(new_border)
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

    def add_location_dvr(self, row_loc, col_loc, style_id, hex_color):
        clean_hex = hex_color.replace("#", "")
        decimal_color = str(int(clean_hex, 16))
        dvr_bucket = self.root.find(".//dataValidationRules")
        if dvr_bucket is None: return
            
        rule = ET.SubElement(dvr_bucket, "dataValidationRule", position="1", name="Auto Format Rule", enabled="true", customStyle="true", rowLocation=str(float(row_loc)), colLocation=str(float(col_loc)))
        cond = ET.SubElement(rule, "dataValidationCond", toolTip="", groupOpenNestingLevel="0", operator="0", displayMessageInDVPane="false", honorPmRules="false", negate="false", groupCloseNestingLevel="0", position="1", styleId=str(style_id), type="8", bgColor=decimal_color, Valid="true", logicalOperator="0")
        ET.SubElement(cond, "compareValue", type="6", value="")
        ET.SubElement(cond, "compareToValue", type="0", value="")

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