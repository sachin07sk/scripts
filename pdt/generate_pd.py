import os

def generate_suite():
    # ── 1. FLOORPLAN REPORTS ──
    with open("innovus_fp.rpt", "w") as f:
        f.write("""=======================================================
Floorplan Design Report (Baseline)
=======================================================
Die Area            : 640000
Core Area           : 512000
Utilization         : 72.5%
Aspect Ratio        : 1.0
""")

    with open("innovus_fp_v2.rpt", "w") as f:
        f.write("""=======================================================
Floorplan Design Report (Version 2 - High Density)
=======================================================
Die Area            : 580000
Core Area           : 464000
Utilization         : 78.2%
Aspect Ratio        : 1.2
""")

    # ── 2. PLACEMENT REPORTS ──
    with open("place.rpt", "w") as f:
        f.write("""=======================================================
Cell Placement Report (Clean Run)
=======================================================
Total Cells         : 12400
Placed Cells        : 12400
Unplaced Cells      : 0
Placement Density   : 72.5
""")

    with open("place_v2.rpt", "w") as f:
        f.write("""=======================================================
Cell Placement Report (Version 2 - Displaced Gates)
=======================================================
Total Cells         : 12100
Placed Cells        : 12050
Unplaced Cells      : 50
Placement Density   : 70.2
""")

    # ── 3. ROUTING REPORTS ──
    with open("route.rpt", "w") as f:
        f.write("""=======================================================
Global & Detail Routing Log (100% Routed)
=======================================================
Total Nets          : 9800
Routed Nets         : 9800
Unrouted Nets       : 0
Total Wire Length   : 185000.5
Via Count           : 48200
""")

    with open("route_v2.rpt", "w") as f:
        f.write("""=======================================================
Global & Detail Routing Log (Version 2 - Open Nets)
=======================================================
Total Nets          : 9600
Routed Nets         : 9580
Unrouted Nets       : 20
Total Wire Length   : 182000
Via Count           : 47500
""")

    # ── 4. DRC & CONGESTION REPORTS ──
    with open("drc.rpt", "w") as f:
        f.write("""=======================================================
Signoff Physical Verification Report (Clean)
=======================================================
Total DRC Violations: 0
Short Violations    : 0
Spacing Violations  : 0
Antenna Violations  : 0

Total Overflow      : 0
Horizontal Overflow : 0
Vertical Overflow   : 0
""")

    with open("drc_v2.rpt", "w") as f:
        f.write("""=======================================================
Signoff Physical Verification Report (Violations Found)
=======================================================
Total DRC Violations: 12
Short Violations    : 3
Spacing Violations  : 7
Antenna Violations  : 2

Total Overflow      : 4
Horizontal Overflow : 2
Vertical Overflow   : 2
""")

    print("Successfully generated 8 Physical Design test files:")
    print("  -> Floorplan : innovus_fp.rpt, innovus_fp_v2.rpt")
    print("  -> Placement : place.rpt, place_v2.rpt")
    print("  -> Routing   : route.rpt, route_v2.rpt")
    print("  -> Signoff   : drc.rpt, drc_v2.rpt")

if __name__ == "__main__":
    generate_suite()
