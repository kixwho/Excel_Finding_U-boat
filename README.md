<img width="1069" height="717" alt="image" src="https://github.com/user-attachments/assets/a041a803-2dce-48d5-b769-f34dbefcb196" />


**Problem**

You are using an underwater probe to search for a sunken U-boat. At any time in the search, your probe is located at some point (x,y) in a grid, where the distance between lines in the grid is some convenient unit such as 100 meters. The U-boat is at some unknown location on the grid, (X,Y).

If your probe is at (x,y), you can only move it to one of the eight nearby grid points (x-1, y-1), (x-1, y), (x-1,y+1), (x, y-1), (x, y+1), (x+1,y-1), (x+1,y), or (x+1,y+1), with probability 1/8 each, for the next search.

If you start at (0,0) and the U-boat is at (5,2), use simulation to estimate the probability that you will find the U-boat in 100 moves or fewer.

**Solution**

Model setup:
1. •	A lookup table of all possible moves
2. •	Use U-boat location (5,2)
3. •	Start each trial at (0,0)
4.    For each iteration:
  
   •	Randomly pick one row from the lookup (uniform), add its x/y changes to current position;
   
   •	Repeat 100 times

   •	If at any move the position reaches U-boat's location, that iteration is a success.
   
6. •	Repeat for 10,000 iterations and recorded how many iterations had success.

**Simulation**

There are a few different ways to formulate and solve this problem. For the simulation, my first attempt was done using @Risk, a commercial Monte Carlo engine.

Since the model is pretty straightforward, I've also tried running the sim using VBA.

The results are the same, we manage to find the U-boat within 100 moves **16%** of the time. The macro did run a lot faster than @Risk, though, you can find a scatter plot of where we eventually end up in the Excel file. 🚢⚓
