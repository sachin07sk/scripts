import os

# 1. SystemVerilog ALU Testbench Code
ALU_CODE = """
module tb_alu;
  // Mock design variables
  reg [3:0] a, b;
  wire [7:0] y;
  
  initial begin
    $display("Starting ALU Regression Test...");
    #10;
    // Your script parses keywords from the generated stdout redirect log
    $display("Simulation complete. Status: PASS");
    $display("TEST PASSED");
  end
endmodule
"""

# 2. Verilog FIFO Testbench Code
FIFO_CODE = """
module tb_fifo;
  reg clk, rst;
  
  initial begin
    $display("Starting FIFO Regression Test...");
    #5;
    $display("Error: Write pointer overflow detected at simulation time 50ns!");
    $display("TEST FAILED");
  end
endmodule
"""

# 3. A Mock Simulator to replicate Questa execution times and logs
MOCK_SIMULATOR_CODE = """import sys
import time

args = " ".join(sys.argv)

if "test_alu.sv" in args:
    time.sleep(12)  # Matches your target runtime of 12 seconds
    print("Reading design files...")
    print("Optimization complete.")
    print("Simulation complete. Status: PASS")
    print("TEST PASSED")
elif "test_fifo.v" in args:
    time.sleep(8)   # Matches your target runtime of 8 seconds
    print("Reading design files...")
    print("Error: Write pointer overflow detected!")
    print("TEST FAILED")
else:
    print("Simulation running...")
    print("Simulation complete.")
"""

def setup_suite():
    # Write mock hardware design/test files
    with open("test_alu.sv", "w") as f:
        f.write(ALU_CODE.strip())
        
    with open("test_fifo.v", "w") as f:
        f.write(FIFO_CODE.strip())
        
    # Write the mock execution engine
    with open("mock_vsim.py", "w") as f:
        f.write(MOCK_SIMULATOR_CODE.strip())

    print("Successfully generated test files:")
    print("  -> test_alu.sv  (Target: PASS, ~12s)")
    print("  -> test_fifo.v (Target: FAIL, ~8s)")
    print("  -> mock_vsim.py (Stands in for Questa/ModelSim toolchain)")

if __name__ == "__main__":
    setup_suite()
