import unittest
import xml.etree.ElementTree as ET
import sys
import types


tkinter = types.ModuleType("tkinter")
tkinter.Tk = type("Tk", (), {})
tkinter.ttk = types.ModuleType("tkinter.ttk")
tkinter.messagebox = types.ModuleType("tkinter.messagebox")
tkinter.filedialog = types.ModuleType("tkinter.filedialog")

sys.modules.setdefault("tkinter", tkinter)
sys.modules.setdefault("tkinter.ttk", tkinter.ttk)
sys.modules.setdefault("tkinter.messagebox", tkinter.messagebox)
sys.modules.setdefault("tkinter.filedialog", tkinter.filedialog)

from sudoc_export import parse_record


def build_record(ind2_attr, field_214, field_210):
    ind2 = "" if ind2_attr is None else f' ind2="{ind2_attr}"'
    xml = f"""
    <record>
      <controlfield tag="001">123456789</controlfield>
      <leader>00000nam a2200000   4500</leader>
      <datafield tag="200"><subfield code="a">Titre</subfield></datafield>
      {field_214.format(ind2=ind2)}
      {field_210}
    </record>
    """
    return ET.fromstring(xml)


class ParseRecord214Tests(unittest.TestCase):
    def test_214_c_used_when_ind2_missing(self):
        record = build_record(
            None,
            """
            <datafield tag="214" ind1=" ">
              <subfield code="a">Paris</subfield>
              <subfield code="c">Éditeur 214</subfield>
              <subfield code="d">2024</subfield>
            </datafield>
            """,
            """
            <datafield tag="210" ind1=" " ind2=" ">
              <subfield code="a">Lyon</subfield>
              <subfield code="c">Éditeur 210</subfield>
              <subfield code="d">1999</subfield>
            </datafield>
            """,
        )

        data = parse_record(record)

        self.assertEqual(data["editeur"], "Éditeur 214")
        self.assertEqual(data["lieu"], "Paris")
        self.assertEqual(data["date"], "2024")

    def test_214_c_used_when_ind2_is_zero(self):
        record = build_record(
            "0",
            """
            <datafield tag="214" ind1=" "{ind2}>
              <subfield code="a">Paris</subfield>
              <subfield code="c">Éditeur 214</subfield>
              <subfield code="d">2024</subfield>
            </datafield>
            """,
            """
            <datafield tag="210" ind1=" " ind2=" ">
              <subfield code="c">Éditeur 210</subfield>
            </datafield>
            """,
        )

        data = parse_record(record)

        self.assertEqual(data["editeur"], "Éditeur 214")

    def test_214_c_used_when_ind2_is_one(self):
        record = build_record(
            "1",
            """
            <datafield tag="214" ind1=" "{ind2}>
              <subfield code="a">Marseille</subfield>
              <subfield code="c">Éditeur 214</subfield>
              <subfield code="d">2023</subfield>
            </datafield>
            """,
            """
            <datafield tag="210" ind1=" " ind2=" ">
              <subfield code="c">Éditeur 210</subfield>
            </datafield>
            """,
        )

        data = parse_record(record)

        self.assertEqual(data["editeur"], "Éditeur 214")
        self.assertEqual(data["lieu"], "Marseille")
        self.assertEqual(data["date"], "2023")

    def test_214_c_ignored_when_ind2_is_other_value(self):
        record = build_record(
            "2",
            """
            <datafield tag="214" ind1=" "{ind2}>
              <subfield code="a">Paris</subfield>
              <subfield code="c">Éditeur 214</subfield>
              <subfield code="d">2024</subfield>
            </datafield>
            """,
            """
            <datafield tag="210" ind1=" " ind2=" ">
              <subfield code="a">Lyon</subfield>
              <subfield code="c">Éditeur 210</subfield>
              <subfield code="d">1999</subfield>
            </datafield>
            """,
        )

        data = parse_record(record)

        self.assertEqual(data["editeur"], "Éditeur 210")
        self.assertEqual(data["lieu"], "Lyon")
        self.assertEqual(data["date"], "1999")

    def test_selected_214_keeps_place_and_date_from_same_field(self):
        record = ET.fromstring(
            """
            <record>
              <controlfield tag="001">123456789</controlfield>
              <leader>00000nam a2200000   4500</leader>
              <datafield tag="200"><subfield code="a">Titre</subfield></datafield>
              <datafield tag="214" ind1=" " ind2="2">
                <subfield code="a">Paris</subfield>
                <subfield code="c">Éditeur refusé</subfield>
                <subfield code="d">2024</subfield>
              </datafield>
              <datafield tag="214" ind1=" " ind2="1">
                <subfield code="a">Marseille</subfield>
                <subfield code="c">Éditeur retenu</subfield>
                <subfield code="d">2023</subfield>
              </datafield>
              <datafield tag="210" ind1=" " ind2=" ">
                <subfield code="a">Lyon</subfield>
                <subfield code="c">Éditeur 210</subfield>
                <subfield code="d">1999</subfield>
              </datafield>
            </record>
            """
        )

        data = parse_record(record)

        self.assertEqual(data["editeur"], "Éditeur retenu")
        self.assertEqual(data["lieu"], "Marseille")
        self.assertEqual(data["date"], "2023")


if __name__ == "__main__":
    unittest.main()
