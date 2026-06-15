import os

def generate_test_scripts():
    # 1. TCL Jobs (PASS Cases)
    tcl_pass_script = """# Dummy TCL Job Script
puts "Starting VLSI Task..."
after 2000 ;# Simulates execution delay
puts "Task completed successfully."
exit 0
"""
    
    tcl_jobs = ["synth_block1.tcl", "synth_block2.tcl", "pnr_block1.tcl", "sta_signoff.tcl"]
    for job in tcl_jobs:
        with open(job, "w") as f:
            f.write(tcl_pass_script)

    # 2. Shell Script Job (FAIL Case)
    sh_fail_script = """#!/bin/bash
echo "Starting Functional Verification..."
sleep 3
echo "Error: U_ALU mismatch detected at cycle 450!"
exit 1
"""
    with open("sim_test_alu.sh", "w") as f:
        f.write(sh_fail_script)

    # 3. Shell Script Job (PASS Case)
    sh_pass_script = """#!/bin/bash
echo "Starting Functional Verification..."
sleep 2
echo "All tests matched successfully."
exit 0
"""
    with open("sim_test_fifo.sh", "w") as f:
        f.write(sh_pass_script)

    print("Successfully generated 6 distinct VLSI test job scripts:")
    print("  -> synth_block1.tcl  (PASS)")
    print("  -> synth_block2.tcl  (PASS)")
    print("  -> sim_test_alu.sh   (FAIL)")
    print("  -> sim_test_fifo.sh  (PASS)")
    print("  -> pnr_block1.tcl    (PASS)")
    print("  -> sta_signoff.tcl   (PASS)")

if __name__ == "__main__":
    generate_test_scripts()
