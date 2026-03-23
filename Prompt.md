I want you to write a Python script that builds a daily forward curve using linear interpolation from an Excel input file.

Task:
The input Excel file has exactly two columns:
1. days_to_expiry
2. fwd_rate

Each row is one known point on the curve. For example, days_to_expiry may contain values like 7, 30, 180, 365, and fwd_rate is the corresponding forward rate.

What I need:
Read the Excel file, sort the data by days_to_expiry ascending, and generate a full daily curve for every integer day between the smallest and largest days_to_expiry in the input.
Use piecewise linear interpolation on fwd_rate with respect to days_to_expiry.
Do not extrapolate beyond the minimum or maximum input day.
For input days that already exist, keep the original value unchanged.
For missing integer days in between, compute the interpolated forward rate.

Output:
Create a new Excel file containing:
1. days_to_expiry
2. fwd_rate_interpolated

The output should include every integer day from the minimum input day to the maximum input day, inclusive.

Requirements:
- Use Python
- Use pandas and numpy only unless there is a strong reason otherwise
- Write clean, readable code
- Add basic validation:
  - confirm the required two columns exist
  - drop rows with missing values in these two columns
  - ensure days_to_expiry is numeric
  - if duplicate days_to_expiry exist, keep the last one after sorting, unless you think a better approach is more robust, but explain it
- The script should be easy to run locally
- Put the input file path and output file path near the top of the script as user-editable variables
- Add concise comments explaining the logic
- After writing the script, also provide a short explanation of how it works

Implementation details:
- Build a complete integer grid using all whole-number days from min(days_to_expiry) to max(days_to_expiry)
- Merge or reindex the original known points onto that grid
- Perform linear interpolation based on days_to_expiry
- Save the final result to Excel

Please produce:
1. the full Python script
2. a brief explanation
3. any assumptions you are making