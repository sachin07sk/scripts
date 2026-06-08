import os

TEMPLATE = """========================================================
Verification Coverage Report
========================================================

--- Global Metrics ---
Line Coverage      : {line}%
Branch Coverage    : {branch}%
Condition Coverage : {condition}%
Toggle Coverage    : {toggle}%
FSM Coverage       : {fsm}%
Functional Coverage: {functional}%

--- Functional Bin Summary ---
Total Bins         : {total_bins}
Covered Bins       : {covered_bins}
Uncovered Bins     : {uncovered_bins}

--- Hierarchical Module Breakdown ---
{modules_data}
========================================================
"""

def generate_files():
    # 1. Block 1: High Coverage Report
    block1_modules = ""
    with open("cov_block1.log", "w") as f:
        f.write(TEMPLATE.format(
            line="98.5", branch="95.2", condition="91.0", 
            toggle="96.8", fsm="100.0", functional="93.5",
            total_bins="4500", covered_bins="4432", uncovered_bins="68",
            modules_data=block1_modules
        ))

    # 2. Block 2: Low Coverage Report (Triggers Red/Yellow warning highlights)
    block2_modules = ""
    with open("cov_block2.log", "w") as f:
        f.write(TEMPLATE.format(
            line="82.3", branch="78.5", condition="70.2", 
            toggle="85.0", fsm="90.0", functional="80.1",
            total_bins="3200", covered_bins="2624", uncovered_bins="576",
            modules_data=block2_modules
        ))

    # 3. Top Chip: Comprehensive Full System Coverage Report + Module Breakdown
    top_modules = [
        "alu_top       98.5%",
        "fifo_ctrl     96.2%",
        "uart_tx       88.0%",
        "spi_master    82.5%",
        "clk_div       100.0%"
    ]
    with open("cov_top.log", "w") as f:
        f.write(TEMPLATE.format(
            line="96.1", branch="93.4", condition="88.5", 
            toggle="94.2", fsm="98.5", functional="91.0",
            total_bins="8000", covered_bins="7680", uncovered_bins="320",
            modules_data="\n".join(top_modules)
        ))

    print("Successfully generated 3 Coverage .log test files:")
    print("  -> cov_block1.log (High coverage metrics)")
    print("  -> cov_block2.log (Low coverage metrics triggering color conditional fills)")
    print("  -> cov_top.log    (Top-level chip matrix with 5 design modules)")

if __name__ == "__main__":
    generate_files()
