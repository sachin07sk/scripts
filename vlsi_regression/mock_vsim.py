import sys
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