import os

TEMPLATE = """========================================================
Synthesis & Implementation QoR Report
========================================================

--- Timing Summary ---
WNS             : {wns}
TNS             : {tns}
Violated Paths  : {viols}

--- Power Summary ---
Total Power     : {tot_pwr}
Dynamic Power   : {dyn_pwr}
Leakage Power   : {leak_pwr}

--- Area Summary ---
Cell Count      : {cells}
Total Area      : {area}
Utilization     : {util}%
========================================================
"""

def generate_files():
    # 1. Baseline Report
    with open("baseline_qor.rpt", "w") as f:
        f.write(TEMPLATE.format(
            wns="0.12", tns="0.00", viols="0",
            tot_pwr="55.3", dyn_pwr="45.0", leak_pwr="10.3",
            cells="4500", area="12400", util="72.0"
        ))

    # 2. Run2 Report
    with open("run2_qor.rpt", "w") as f:
        f.write(TEMPLATE.format(
            wns="-0.05", tns="-0.23", viols="12",
            tot_pwr="53.1", dyn_pwr="43.1", leak_pwr="10.0",
            cells="4400", area="12100", util="70.5"
        ))

    # 3. Run3 Report
    with open("run3_qor.rpt", "w") as f:
        f.write(TEMPLATE.format(
            wns="-0.31", tns="-1.20", viols="45",
            tot_pwr="48.9", dyn_pwr="39.9", leak_pwr="9.0",
            cells="4400", area="12100", util="70.5"
        ))

    print("Successfully generated 3 QoR .rpt files!")

if __name__ == "__main__":
    generate_files()
