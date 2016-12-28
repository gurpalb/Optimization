
This Solvers folder, and the appropriate 32 or 64 bit sub-folder (cjhosen to match the operating system bit-ness), is added to SolverStudio's PATH environment variable when running PuLP, and when launching any of the solvers such as AMPL.

This Solvers directory is added to the end of the PATH, and one of Solvers/32bit or Solvers/64bit is also added, depending on the bit-ness of the operating system. (Note that this matches the operating system bitness, not the process bitness, as these are stand-alone solvers, not DLLs.) Thus, these solvers are used if no other versions are installed on the machine.

When running the student version of AMPL, the AMPLCML folder as downloaded from AMPL is placed at the start of the path, ensuring the license-free solvers in this folder are used in preference to any installed solvers (which may require and consume licenses).

You should place solvers that you only want to run on a particular operating system bitness in the appropriate sub-folder, eg Solvers/32bit. A solver that should run under both bitnesses should go in Solvers.

