# Task 4 TrackBy Evidence

Existing stable `trackBy` usage was preserved:
- Company duplicate and normal rows use `trackByKey`.
- Product modal rows use `trackByKey`.
- Price company groups, price rows, and price groups use `trackByKey`.
- Misorder groups/items use `trackByKey`.
- Near-phrase groups use `trackByKey`.
- Repeated phrase rows use `trackByIndex` because they are editable positional form rows.

No additional `trackBy` change was needed for the confirmed custom-configuration lag after Task 1.
