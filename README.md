# Beam Formulae functions
Functions for returning moment, shear and deflection for a simply supported beam with or without a cantilever

Refer to this series of blog posts for details on use of these functions and check out the example in the Beam Formulae excel file, enjoy!
https://engineervsheep.com/2021/moment-shear-deflection-functions-1/

### Change log

2021-02-03 - Initial commit

2021-02-28 - Remove ThisWorkbook.cls from commit

2021-02-28 - Fix cantilever functions to fix error regarding deflection including TAN() of slopes/rotations (small error in calculated deflections resulted, becoming more prominant with larger deflections>100mm). Initially differences put down to numerical precision issues

2021-02-28 - Add formulation for considering point moments on main and cantilever span
