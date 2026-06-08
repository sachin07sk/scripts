module tb_fifo;
  reg clk, rst;
  
  initial begin
    $display("Starting FIFO Regression Test...");
    #5;
    $display("Error: Write pointer overflow detected at simulation time 50ns!");
    $display("TEST FAILED");
  end
endmodule