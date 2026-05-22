# Task 6 Browser Regression Evidence

Post-fix timings on the same fixture:
- Upload file to Step 2: 12,660 ms
- Check companies to Step 3: 18,382 ms
- Profile select Sơn Phương -> Cao Thành: primary acknowledgement 119 ms
- Profile select Cao Thành -> Sơn Phương: primary acknowledgement 117 ms
- Verify prefixes click: 2,523 ms

Baseline comparison:
- Profile select Sơn Phương -> Cao Thành baseline: 2,933 ms, now primary acknowledgement 119 ms.
- Profile select Cao Thành -> Sơn Phương baseline: 1,588 ms, now primary acknowledgement 117 ms.
- Verify prefixes baseline: 3,515 ms, now 2,523 ms.

Delayed loading:
- Slow save-config API response delayed to 2300 ms.
- Loading indicator hidden at 1500 ms, visible at 2250 ms, and cleared after completion.
