package docxpatch

import (
	"encoding/xml"
	"strconv"
)

// NativeDocxApproximatedSettingV1 records source facts explicitly disregarded by
// the read-only current-layout policy, never by strict pagination or mutation.
type NativeDocxApproximatedSettingV1 struct {
	Kind   string            `json:"kind"`
	Path   string            `json:"path"`
	Values map[string]string `json:"values"`
}

func nativeApproximationSettingReason(fact NativeDocxApproximatedSettingV1) string {
	return "Current-layout approximation disregards " + fact.Kind + " at " + fact.Path + "; source values are retained and Word layout may differ"
}

func nativeApproximateSetting(node *nativeXMLNode, wordNS string) *NativeDocxApproximatedSettingV1 {
	result := &NativeDocxApproximatedSettingV1{Kind: node.Name.Local, Path: node.Path, Values: map[string]string{}}
	if node.Name.Space == wordNS {
		switch node.Name.Local {
		case "themeFontLang":
			if !nativeExactLeaf(node, xml.Name{Space: wordNS, Local: "val"}, xml.Name{Space: wordNS, Local: "eastAsia"}, xml.Name{Space: wordNS, Local: "bidi"}) {
				return nil
			}
			for _, key := range []string{"val", "eastAsia", "bidi"} {
				if value, present := nativeAttr(node, wordNS, key); present {
					if (key == "val" && value != "en-US" && value != "fr-FR") || (key != "val" && value != "x-none") {
						return nil
					}
					result.Values[key] = value
				}
			}
			if result.Values["val"] == "" {
				return nil
			}
		case "decimalSymbol", "listSeparator":
			if !nativeExactLeaf(node, xml.Name{Space: wordNS, Local: "val"}) {
				return nil
			}
			value, _ := nativeAttr(node, wordNS, "val")
			if (node.Name.Local == "decimalSymbol" && value != "." && value != ",") || (node.Name.Local == "listSeparator" && value != "," && value != ";") {
				return nil
			}
			result.Values["val"] = value
		case "shapeDefaults":
			const officeNS = "urn:schemas-microsoft-com:office:office"
			const vmlNS = "urn:schemas-microsoft-com:vml"
			if !nativeExactContainer(node) || len(node.Children) != 2 {
				return nil
			}
			defaults, layout := node.Children[0], node.Children[1]
			if defaults.Name != (xml.Name{Space: officeNS, Local: "shapedefaults"}) || !nativeExactLeaf(defaults, xml.Name{Space: vmlNS, Local: "ext"}, xml.Name{Local: "spidmax"}) || layout.Name != (xml.Name{Space: officeNS, Local: "shapelayout"}) || !nativeExactContainer(layout, xml.Name{Space: vmlNS, Local: "ext"}) || len(layout.Children) != 1 {
				return nil
			}
			idmap := layout.Children[0]
			if idmap.Name != (xml.Name{Space: officeNS, Local: "idmap"}) || !nativeExactLeaf(idmap, xml.Name{Space: vmlNS, Local: "ext"}, xml.Name{Local: "data"}) {
				return nil
			}
			for _, child := range []*nativeXMLNode{defaults, layout, idmap} {
				if value, _ := nativeAttr(child, vmlNS, "ext"); value != "edit" {
					return nil
				}
			}
			spid, _ := nativeAttr(defaults, "", "spidmax")
			value, err := strconv.ParseUint(spid, 10, 31)
			if err != nil || value == 0 || strconv.FormatUint(value, 10) != spid {
				return nil
			}
			if data, _ := nativeAttr(idmap, "", "data"); data != "1" {
				return nil
			}
			result.Values["spidmax"], result.Values["idmap"] = spid, "1"
		default:
			return nil
		}
		return result
	}
	if node.Name != (xml.Name{Space: nativeMathNamespace, Local: "mathPr"}) || !nativeExactContainer(node) {
		return nil
	}
	allowed := map[string]string{"mathFont": "Cambria Math", "brkBin": "before", "brkBinSub": "--", "lMargin": "0", "rMargin": "0", "defJc": "centerGroup", "wrapIndent": "1440", "intLim": "subSup", "naryLim": "undOvr"}
	for _, child := range node.Children {
		if child.Name.Space != nativeMathNamespace || !nativeExactLeaf(child, xml.Name{Space: nativeMathNamespace, Local: "val"}) {
			return nil
		}
		key := child.Name.Local
		if _, seen := result.Values[key]; seen {
			return nil
		}
		value, present := nativeAttr(child, nativeMathNamespace, "val")
		if key == "smallFrac" || key == "dispDef" {
			if !present {
				value = "true"
			} else if _, valid := nativeLexicalOnOff(value); !valid {
				return nil
			}
		} else if expected, exists := allowed[key]; !exists || !present || value != expected {
			return nil
		}
		result.Values[key] = value
	}
	if len(result.Values) != 11 {
		return nil
	}
	return result
}
