import os

# --- LEF TEMPLATES ---
LEF_TECH = """LAYER Metal1
  TYPE ROUTING ;
  PITCH 0.2 ;
  WIDTH 0.1 ;
END Metal1
LAYER Via1
  TYPE CUT ;
END Via1
"""

LEF_MACRO = """MACRO {macro_name}
  CLASS CORE ;
  SIZE {w} BY {h} ;
  PIN A
    DIRECTION INPUT ;
    USE SIGNAL ;
    LAYER Metal1 ;
      RECT -0.05 -0.05 0.05 0.05 ;
  END A
  PIN Y
    DIRECTION OUTPUT ;
    USE SIGNAL ;
    LAYER Metal1 ;
      RECT 0.95 0.95 1.05 1.05 ;
  END Y
END {macro_name}
"""

# --- DEF TEMPLATES ---
DEF_TEMPLATE = """DESIGN {design_name} ;
DIEAREA ( 0 0 ) ( 5000 5000 ) ;

COMPONENTS 2 ;
- U1 {macro_name} + PLACED ( 100 200 ) N ;
- U2 {macro_name} + PLACED ( 500 600 ) FN ;
END COMPONENTS

PINS 2 ;
- IN_A + NET IN_A
  + DIRECTION INPUT + USE SIGNAL
  + LAYER Metal1 ( -10 10 ) ( 10 10 ) ;
- OUT_Y + NET OUT_Y
  + DIRECTION OUTPUT + USE SIGNAL
  + LAYER Metal2 ( -10 10 ) ( 10 10 ) ;
END PINS

NETS 2 ;
- IN_A ( PIN IN_A ) ( U1 A ) ;
- n1 ( U1 Y ) ( U2 A ) ;
END NETS
"""

def generate_files():
    # 1. Create 5 LEF files
    with open("tech.lef", "w") as f:
        f.write(LEF_TECH)
    
    for i in range(1, 5):
        with open(f"macro_library_{i}.lef", "w") as f:
            f.write(LEF_MACRO.format(macro_name=f"AND2_X{i}", w=0.8+(i*0.1), h=1.2))

    # 2. Create 5 DEF files
    for i in range(1, 6):
        with open(f"block_{i}.def", "w") as f:
            f.write(DEF_TEMPLATE.format(design_name=f"chip_block_{i}", macro_name=f"AND2_X{1}"))

    print("Successfully generated 5 .lef and 5 .def files in the current directory!")

if __name__ == "__main__":
    generate_files()
