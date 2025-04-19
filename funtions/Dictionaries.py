code_to_gauge = {
    "3000-ALU-0.080": "080 AL 3000",
    "3000-ALU-0.0": "080 AL 3000",
    "3000-ALU-0.125": "125 AL 3000",
    "GALV-0.062": "16 GA GALV",
    "GALV-0.078": "14 GA GALV",
    "GALV-0.140": "10 GA GALV",
    "GALV-0.14": "10 GA GALV",
    "STEEL-0.187": "7 GA",
    "STEEL-0.250": "25 INCH",
    "STEEL-0.25": "25 INCH",
    "GALV-0.102": "12 GA GALV",
    "STEEL-0.313": "312 INCH",
    "DMND-PLT-0.250": "DIA 1/4",
    "304-SS-0.140": "10 GA 304SS",
    "STEEL-0.750": "750 INCH",
    "STEEL-0.500": "500 INCH",
    "304-SS-0.250": "25 INCH 304SS",
    "304-SS-0.187": "7 GA 304SS",
    "STEEL-PERF-0.125": "11 GA PERF",
    "316-SS-0.078": "14 GA 316SS",
    "316-SS-0.062": "16 GA 316SS",
    "STEEL-0.140": "10 GA",
    "STEEL-0.14": "10 GA",
    "STEEL-0.078": "14 GA",
    "GALV-0.102": "12 GA GALV",
    "STEEL-PERF-0.062": "16 GA PERF"
}

categories = {
    "ENCL": "ENCL",
    "ENGR": "ENGR",
    "MEC": "MEC",
    "PARTS ORDER": "PARTS ORDER",
    "REWORK": "REWORK",
    "SIL": "SIL",
    "TANK": "TANK"
}

Material_dict = {
    "1. Steel": {"tiene": ["STEEL"], "no_tiene": ["STEEL-PERF"]},
    "2. Galv": {"tiene": ["GALV"], "no_tiene": []},
    "3. Diamond Plate": {"tiene": ["DMND"], "no_tiene": []},
    "4. Alum": {"tiene": ["3000-ALU"], "no_tiene": []},
    "5. Perf": {"tiene": ["STEEL-PERF"], "no_tiene": []},
    "6. Steel 304ss": {"tiene": ["304-SS"], "no_tiene": []},
    "7. Steel 316ss": {"tiene": ["316-SS"], "no_tiene": []},
    "8. Steel Perf": {"tiene": ["SS-PERF"], "no_tiene": []}
}

valid_sizes = ["60x86","60x120","60x144","72x120","72x144","48x144","60x145","72x145",
               "72x187","60x157","48x90","48x157","48x120","48x158","48x96","23.94x155",
               "23.94x190","23.94x200","33x121","33x145","24.83x133","24.83x170","24.83x190",
               "24.83x205","24.96x145","24.96x200","28x120","28x157","44.24x128","44.24x151",
               "24.45x145","24.45x205","27.92x120","27.92x157","44.21x120","44.21x157","60x175",
               "60x187","60x164","60x172","60x163","45x174","60x162","23.94x205","23.94x191.75",
               "24.83x206.75","24.83x149","33x149","60x146","170X24.833","190X23.95","157X27.93",
               "145X33.015","205X24.8339","151X44.24","121X33.012","133X24.8339","90X48.1",
               "175X50","144X50"]

type_dict = {
    "ENCL": "2.Encl",
    "ENGR": "6.Engr",
    "MEC": "3.Mec",
    "PARTS ORDER": "4.Parts Order",
    "REWORK": "5.Rework",
    "SIL": "4.Sil",
    "TANK": "1.Tank"
}

machine_dict = {
    "Amada_ENSIS_4020AJ": "1. Laser",
    "Messer_170Amp_Plasm": "2.Plasma",
    "Amada_Vipros_358K": "3.Turret",
    "Amada_EMK3612M2": "4.Citurret"
}