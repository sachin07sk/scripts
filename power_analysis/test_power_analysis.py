import os

LOG_TEMPLATE = """===================================================
Questa Power Analysis Report
===================================================
Power Unit       : mW

Dynamic Power    : {dyn}
Internal Power   : {int}
Leakage Power    : {leak}
Total Power      : {tot}

----------------------------------------------------------------------
Instance                            Cell       Dynamic Internal Leakage Total
----------------------------------------------------------------------
{cells}
----------------------------------------------------------------------
Total
"""

def generate_files():
    # 1. Create the specific chip_power.log to match your expected output
    cells_data = []
    # Add the top 2 consumers you expect
    cells_data.append("CLK_BUF_0                           CLKBUF_X4  2.00    1.50     0.73    4.23")
    cells_data.append("CORE/MUL_U1                         MUL32      1.80    1.20     0.91    3.91")
    
    # Pad with 1,248 dummy cells to hit exactly 1,250 total cells
    for i in range(1248):
        cells_data.append(f"DUMMY_INST_{i:<23} DFF_X1     0.01    0.01     0.01    0.03")
        
    chip_power_content = LOG_TEMPLATE.format(
        dyn="45.23", int="12.10", leak="3.45", tot="60.78",
        cells="\n".join(cells_data)
    )
    
    with open("chip_power.log", "w") as f:
        f.write(chip_power_content)

    # 2. Create a few smaller block-level log files for multi-file testing
    for i in range(1, 4):
        generic_cells = []
        for j in range(150):
            generic_cells.append(f"BLOCK_{i}_INST_{j:<21} AND2_X2    0.10    0.05     0.02    0.17")
            
        generic_content = LOG_TEMPLATE.format(
            dyn=f"{15.0 + i:.2f}", 
            int=f"{5.0 + i:.2f}", 
            leak=f"{1.0 + (i*0.5):.2f}", 
            tot=f"{21.0 + (i*2.5):.2f}",
            cells="\n".join(generic_cells)
        )
        with open(f"block_{i}_power.log", "w") as f:
            f.write(generic_content)

    print("Successfully generated 4 .log test files!")

if __name__ == "__main__":
    generate_files()
