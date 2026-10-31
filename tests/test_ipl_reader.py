"""acd.ipl_reader: S1000D IPD data modules and ATA iSpec 2200 IPLs into one neutral parts list."""

from pathlib import Path

import pytest

from acd.ipl_reader import IplReadError, parse_markup, read_ipl, read_ispec2200_ipl, read_s1000d_ipd

SOURCES = Path(r"D:\CMM\CMM_AI_Authoring\01 Documents for AI")
S1000D = SOURCES / "03 S1000D"


def items(data):
    return [it for f in data["figures"] for it in f["items"]]


# -- the tolerant tree


def test_omitted_end_tags_close_at_the_next_start_tag_or_same_name():
    root = parse_markup("<A><B x=1>one<C>two<B>three</A>")
    a = root.child("a")
    assert [c.tag for c in a.children] == ["B", "C", "B"]
    assert a.children[0].get("X") == "1"
    assert [c.own_text() for c in a.children] == ["one", "two", "three"]


def test_empty_elements_hold_nothing_and_closed_names_nest():
    root = parse_markup('<csn><ipp ippn="K"><isn id="00A"><rfs value="0"><qna>REF</qna></isn></csn>')
    csn = root.child("csn")
    assert [c.name for c in csn.children] == ["ipp", "isn"]
    assert csn.find("qna").text() == "REF"
    assert csn.child("isn").path() == "csn/isn"


def test_doctype_subset_entities_pi_cdata_comment_and_marked_sections():
    text = ('<?xml version="1.0"?><!DOCTYPE x [<!ENTITY co "ACME"> <!ENTITY g SYSTEM "g.cgm" NDATA cgm>]>'
            '<?Pub Inc><x a="&amp;&#65;">&co; &deg; &unknown; <!-- c --><![CDATA[<b>]]>'
            '<![IGNORE[<y>gone</y>]]><![INCLUDE[<z>kept</z>]]></x>')
    x = parse_markup(text).child("x")
    assert x.get("a") == "&A"
    assert x.own_text().startswith("ACME \u00b0 &unknown; <b>")
    assert x.find("y") is None
    assert x.find("z").text() == "kept"


def test_empty_end_tag_and_paths_with_indexes():
    root = parse_markup("<r><s>1</><s>2</s><t/></r>")
    r = root.child("r")
    assert [c.text() for c in r.children if c.name == "s"] == ["1", "2"]
    assert r.children[1].path() == "r/s[2]"
    assert r.child("t").path(base=r) == "t"


def test_no_ipl_markup_raises():
    with pytest.raises(IplReadError):
        read_ipl("<dmodule><content><procedure/></content></dmodule>")


# -- S1000D

IPD_42 = """<?xml version="1.0" encoding="UTF-8"?>
<dmodule><identAndStatusSection><dmAddress><dmIdent><dmCode modelIdentCode="ACME" systemDiffCode="A"
 systemCode="32" subSystemCode="1" subSubSystemCode="1" assyCode="11" disassyCode="00"
 disassyCodeVariant="A" infoCode="941" infoCodeVariant="A" itemLocationCode="D"/></dmIdent>
 <dmAddressItems><dmTitle><techName>Actuator</techName><infoName>Illustrated parts data</infoName></dmTitle>
 </dmAddressItems></dmAddress></identAndStatusSection>
<content><referencedApplicGroup><applic id="app-0001"><displayText><simplePara>Pre SB 32-001</simplePara>
</displayText></applic></referencedApplicGroup>
<illustratedPartsCatalog>
 <figure id="fig-1"><title>Actuator</title><graphic infoEntityIdent="ICN-1"/><graphic infoEntityIdent="ICN-2"/></figure>
 <catalogSeqNumber indenture="1" figureNumber="01" figureNumberVariant="A" item="000">
  <itemSeqNumber itemSeqNumberValue="00A"><quantityPerNextHigherAssy>REF</quantityPerNextHigherAssy>
   <partRef manufacturerCodeValue="D9893" partNumberValue="2681A0000-01"/>
   <partSegment><itemIdentData><descrForPart>Actuator, rotary</descrForPart></itemIdentData></partSegment>
  </itemSeqNumber></catalogSeqNumber>
 <catalogSeqNumber indenture="2" figureNumber="01" item="010">
  <itemSeqNumber itemSeqNumberValue="00A"><quantityPerNextHigherAssy>1</quantityPerNextHigherAssy>
   <partRef manufacturerCodeValue="D9893" partNumberValue="P-10"/>
   <partSegment><itemIdentData><descrForPart>Housing</descrForPart></itemIdentData></partSegment>
   <partLocationSegment><notIllustrated/><descrForLocation>LH</descrForLocation>
    <catalogSeqNumberRef figureNumber="02" item="001"/></partLocationSegment>
   <applicabilitySegment><usableOnCodeAssy>A</usableOnCodeAssy></applicabilitySegment>
  </itemSeqNumber>
  <itemSeqNumber itemSeqNumberValue="05A" applicRefId="app-0001"><quantityPerNextHigherAssy>1</quantityPerNextHigherAssy>
   <partRef manufacturerCodeValue="D9893" partNumberValue="P-10B"/>
  </itemSeqNumber></catalogSeqNumber>
 <catalogSeqNumber indenture="3" figureNumber="01" item="020">
  <itemSeqNumber itemSeqNumberValue="00A"><quantityPerNextHigherAssy>4</quantityPerNextHigherAssy>
   <partRef manufacturerCodeValue="96906" partNumberValue="MS-1"/>
   <partSegment><itemIdentData><descrForPart>Screw</descrForPart></itemIdentData></partSegment>
   <partLocationSegment><attachStoreShipPart attachStoreShipPartCode="1"/>
    <functionalItemRef functionalItemNumber="12M"/></partLocationSegment>
  </itemSeqNumber></catalogSeqNumber>
</illustratedPartsCatalog></content></dmodule>"""


def test_s1000d_42_items_fields_and_paths():
    d = read_ipl(IPD_42)
    assert d["format"] == "s1000d-ipd"
    assert d["document"] == {"dmc": "DMC-ACME-A-32-11-11-00A-941A-D", "title": "Actuator - Illustrated parts data"}
    [f] = d["figures"]
    assert (f["number"], f["title"], f["sheets"], f["at"]) == ("1", "Actuator", 2, ["figure"])
    got = [(i["item"], i["indenture"], i["part_number"], i["nomenclature"], i["quantity"], i["manufacturer"])
           for i in f["items"]]
    assert got == [("0", 1, "2681A0000-01", "Actuator, rotary", "REF", "D9893"),
                   ("10", 2, "P-10", "Housing", "1", "D9893"),
                   ("10A", 2, "P-10B", "", "1", "D9893"),
                   ("20", 3, "MS-1", "Screw", "4", "96906")]
    housing, alt, screw = f["items"][1:]
    assert housing["illustrated"] is False
    assert housing["remarks"] == ["LH", "SEE FIG 2 ITEM 1"]
    assert housing["effectivity"] == ["A"]
    assert housing["at"] == "catalogSeqNumber[2]/itemSeqNumber[1]"
    assert alt["remarks"] == ["APPLIC Pre SB 32-001"]
    assert screw["attaching"] is True and screw["remarks"] == ["FIN 12M"]
    msgs = " / ".join(n["message"] for n in d["notes"])
    assert "2 item sequence numbers (00A, 05A)" in msgs
    assert "1 parts without descrForPart" in msgs and "P-10B" in msgs
    assert "applicability app-0001 (Pre SB 32-001)" in msgs


def test_s1000d_figure_variants_count_only_when_two_differ():
    one = read_s1000d_ipd(IPD_42)
    assert [f["number"] for f in one["figures"]] == ["1"]
    two = IPD_42.replace('figureNumber="01" item="020"', 'figureNumber="01" figureNumberVariant="B" item="020"')
    assert [f["number"] for f in read_s1000d_ipd(two)["figures"]] == ["1A", "1", "1B"]


IPD_23_SGML = """<!DOCTYPE dmodule PUBLIC "-//S1000D//DTD S1000D IPD 20070228//EN" [
<!ENTITY ICN-1 SYSTEM "ICN-1.CGM" NDATA cgm>
]>
<dmodule><content><ipc>
<figure id ="fig-0001"><title>Light system</title><graphic boardno ="ICN-1"></graphic></figure>
<ipp ippn ="KZ7771111" ips ="LIGHT">
<csn csn ="D00000002A000 " ind ="1" item ="000 "><isn isn ="00A"><rfs value ="0">
<qna>REF</qna><mfc>KZ777</mfc><pnr>LRU1001</pnr><pas><dfp>Light system</dfp></pas>
</isn></csn>
<csn csn ="D00000002A001 " ind ="2"><isn isn ="00A"><rfs value ="1">
<qna>2</qna><mfc>KZ777</mfc><pnr>LRU1010</pnr><pas><dfp>Light, front</dfp></pas><cbs><asp asp="1"></cbs>
</isn></csn>
</ipc></content></dmodule>"""


def test_s1000d_23_sgml_csn_values():
    d = read_ipl(IPD_23_SGML)
    [f] = d["figures"]
    assert f["number"] == "2" and f["title"] == "Light system"
    assert [(i["item"], i["indenture"], i["part_number"], i["quantity"], i["attaching"]) for i in f["items"]] == [
        ("0", 1, "LRU1001", "REF", False), ("1", 2, "LRU1010", "2", True)]
    assert f["items"][1]["at"] == "csn[2]/isn"


def test_s1000d_zero_figure_becomes_1_with_one_note():
    d = read_ipl(IPD_23_SGML.replace("D00000002A", "D00000000A"))
    assert [f["number"] for f in d["figures"]] == ["1"]
    assert sum("is zero" in n["message"] for n in d["notes"]) == 1


# -- iSpec 2200

ISPEC = """<!DOCTYPE CMM PUBLIC "-//ATA-TEXT//DTD CMM-VER3-LEVEL2//EN" [
<!ENTITY lt SDATA "[lt    ]">
]>
<?Pub Inc>
<cmm chapnbr="27" sectnbr="50" subjnbr="74" spl="D9893"><title>Component Maintenance Manual</title>
<ipl><iplintro><vendlist><vendata><mfr>D9893</mfr><mad>LIEBHERR-AEROSPACE LINDENBERG GMBH, P.O. Box 1363,
88153 Lindenberg, Germany</mad></vendata><vendata><mfr>96906</mfr><mad>MIL STANDARDS</mad></vendata></vendlist></iplintro>
<dplist><figure fignbr="1" confltr="A" varnbr="01"><title>Actuator</title>
<graphic><title>Actuator</title><sheet gnbr="G1" sheetnbr="01"></sheet><sheet gnbr="G2" sheetnbr="02"></sheet></graphic>
<prtlist>
<itemdata itemnbr="001A" indent="1" illusind="0"><pnr>2681A0000-01</pnr><iplnom><nom><kwd>ACTUATOR</kwd>
<adt>ROTARY 59S ASSY</adt></nom><mfr>D9893</mfr></iplnom><effcode>A</effcode><upa>RF</upa></itemdata>
<itemdata itemnbr="010" indent="2"><pnr>MS24665-450</pnr><iplnom><nom><kwd>PIN</kwd><adt>PIN-COTTER</adt></nom>
<mfr>96906</mfr><optmfr><opt>NAS1726</opt><mfr>80205</mfr></optmfr><uoa>used on 001A</uoa></iplnom>
<effcode>A</effcode><effcode>B</effcode><upa>002</upa><sym side="LH"></itemdata>
<itemdata itemnbr="020" indent="3" attach="1"><pnr>NAS-20</pnr><iplnom><nom><kwd>SCREW</kwd></nom></iplnom><upa>4</upa></itemdata>
<itemdata itemnbr="030" indent="2" delitem="1"><pnr>OLD-30</pnr><iplnom><nom><kwd>WASHER</kwd></nom></iplnom><upa>1</upa></itemdata>
</prtlist></figure>
<figure fignbr="2" confltr="A"><title>Valve</title><graphic><sheet gnbr="G3"></sheet></graphic><prtlist>
<itemdata itemnbr="-1"><pnr>V-1</pnr><iplnom><nom><kwd>VALVE</kwd></nom></iplnom><upa>1</upa></itemdata>
</prtlist></figure>
<figure fignbr="2" confltr="B"><title>Valve, variant</title><graphic><sheet gnbr="G4"></sheet></graphic><prtlist>
<itemdata itemnbr="1"><pnr></pnr><iplnom><nom><kwd></kwd></nom></iplnom></itemdata>
</prtlist></figure>
</dplist></ipl></cmm>"""


def test_ispec_figures_items_and_vendors():
    d = read_ipl(ISPEC)
    assert d["format"] == "ispec2200"
    assert d["document"]["chapnbr"] == "27" and d["document"]["title"] == "Component Maintenance Manual"
    assert [(f["number"], f["title"], f["sheets"]) for f in d["figures"]] == [
        ("1", "Actuator", 2), ("2A", "Valve", 1), ("2B", "Valve, variant", 1)]
    top, pin, screw, washer = d["figures"][0]["items"]
    assert (top["item"], top["indenture"], top["illustrated"], top["nomenclature"], top["manufacturer"],
            top["effectivity"], top["quantity"]) == ("1A", 1, False, "ACTUATOR, ROTARY 59S ASSY", "D9893", ["A"], "RF")
    assert pin["nomenclature"] == "PIN-COTTER"  # the additional text already starts with the keyword
    assert pin["remarks"] == ["OPT TO NAS1726 80205", "used on 001A"]
    assert pin["effectivity"] == ["A", "B"] and pin["quantity"] == "002"
    assert pin["at"] == "dplist/figure[1]/prtlist/itemdata[2]"
    assert screw["attaching"] is True and screw["nomenclature"] == "SCREW"
    assert washer["remarks"] == ["DELETED"]
    [valve] = d["figures"][1]["items"]
    assert valve["item"] == "1" and valve["illustrated"] is False
    assert d["figures"][2]["items"] == []
    assert d["vendors"][0] == {"code": "D9893", "name": "LIEBHERR-AEROSPACE LINDENBERG GMBH",
                               "address": "P.O. Box 1363, 88153 Lindenberg, Germany"}
    msgs = " / ".join(n["message"] for n in d["notes"])
    assert "deleted item" in msgs and "figure 2B: no item with a part number" in msgs


def test_ispec_upper_case_sgml_without_end_tags_is_read_alike():
    sgml = ("<DPLIST><FIGURE FIGNBR=\"03\"><TITLE>Cover</TITLE><PRTLIST>"
            "<ITEMDATA ITEMNBR=\"010\" INDENT=\"2\"><PNR>C-1<IPLNOM><NOM><KWD>COVER</KWD></NOM></IPLNOM><UPA>1"
            "</ITEMDATA></PRTLIST></FIGURE></DPLIST>")
    d = read_ispec2200_ipl(sgml)
    [it] = d["figures"][0]["items"]
    assert d["figures"][0]["number"] == "3"
    assert (it["item"], it["part_number"], it["nomenclature"], it["quantity"]) == ("10", "C-1", "COVER", "1")


# -- the real samples of the sources (ALTHOM-CMM-AUTHORING, D1), when they are there

REAL = [
    (S1000D / "PRE 3.0/S1000D Issue 1.7/SGML DTD pacakge/AECMA1-7/DM/SGML/AVEE_IPD.SGM", 12, "RA1M240AAA"),
    (S1000D / "PRE 3.0/S1000D Issue 2.3/Bike data set/samples/sgml/dm/bike/DMC-S1000DLIGHTING-AAA-D00-00-00-00AA-941A-D_003-00.SGM", 20, "LRU1001"),  # noqa: E501
    (S1000D / "POST 4.0/S1000D Issue 4.0/Bike data set/S1000D_4-0_20080919_samples/S1000D_4-0/samples/DMC-S1000DLIGHTING-AAA-D00-00-00-00AA-941A-D_006-00_EN-US.xml", 20, "LRU1001"),  # noqa: E501
    (S1000D / "POST 4.0/S1000D Issue 4.2/Bike Sample Data Set/samples_20170115/DMC-S1000DBIKE-AAA-D00-00-00-01AA-941A-D_008-00_en-US.XML", 11, "BICYCLE-001"),  # noqa: E501
    (S1000D / "POST 4.0/S1000D Issue 6/S1000D Issue 6 Bike Sample Data Set/Bike Data Set for Release number 6 R2/DMC-S1000DLIGHTING-AAA-D00-00-00-01AA-941A-D_010-00_EN-US.XML", 24, "LRU1001"),  # noqa: E501
]


@pytest.mark.parametrize("path,count,first", REAL, ids=["1.7-sgml", "2.3-sgml", "4.0", "4.2", "6"])
def test_real_s1000d_ipd_samples(path, count, first):
    if not path.exists():
        pytest.skip("source folder not here")
    d = read_ipl(path)
    got = items(d)
    assert d["format"] == "s1000d-ipd" and len(d["figures"]) == 1
    assert len(got) == count and got[0]["part_number"] == first
    assert all(i["indenture"] for i in got) and all(i["quantity"] for i in got)
    assert not [n for n in d["notes"] if n["level"] == "error"]


def test_real_ispec_bookplan_template():
    path = SOURCES / "01 General/02_CMM Specifications and Standards/05_SGM and XML/Bookplan_r12.sgm"
    if not path.exists():
        pytest.skip("source folder not here")
    d = read_ipl(path)
    assert d["format"] == "ispec2200"
    assert [f["number"] for f in d["figures"]] == ["1", "2", "3"] and not items(d)
    assert d["vendors"][0]["code"] == "D9893"
    xml = read_ipl(path.with_suffix(".xml"))
    assert [f["number"] for f in xml["figures"]] == ["1", "2", "3"]
