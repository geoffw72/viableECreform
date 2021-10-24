README for the Python codes for "Politically Viable U.S. Electoral College Reform"

Refer to the paper's Appendix section A.5.1, Table A2 for the general flow of these codes.
In this folder, I have included the following codes:

WrongWinner.py
WrongWinnerRegional.py
WrongWinnerCovar.py
Disputed.py
DisputedGamma.py
StatePower.py
GammaCalcVustu

For the basic "WrongWinner" code, there is no separate code for computing partisan bias; 
the wrong-winner codes include a code block to compute the partisan bias.
However, for simulated Gamma distributions of state outcomes,
  it was easier to generate Figure 9's scatter plots in a separate code vs. the disputation sims.

Each code pulls in the Excel input file appropriate for the national landscape under study,
except the "Gamma" files, which generate gamma distributions from user-input k and theta.

All codes are stand-alone programs.
The user is prompted for some key inputs, with "hints" (prompts) for usual values.

The output of each simulation run is an Excel file.
All input and output Excel files are organized in this github repository.
