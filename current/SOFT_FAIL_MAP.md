> 新增：同一WT依Config的Pattern順序，前序終點FBC=0才進入下一Pattern；例如Scan Pass才計March。Map不讀Address Count。詳見PATTERN_SEQUENCE.md。

# Soft Fail BLK Map

- Config D/E/F remains Function TNO / Func. / Test Voltage.
- Enter `Hard Fail` in Config G to mark an exclusion TNO (case insensitive; multiple TNOs supported). It does not create a Pattern Map.
- Other nonempty G cells remain Pattern names and identify that Pattern's cumulative endpoint TNO.
- All configured Function TNOs in Config D2:D901 participate in preceding-WT checks, including rows without a Stop Sign. Monitor TNOs are excluded.
- Each Chip/BLK is evaluated independently. Exclude any BLK with FBC > 0 at a Hard Fail TNO in the current or an earlier WT. Exclusion persists through later WTs.
- WT2 additionally requires FBC = 0 at EVERY WT1 Function TNO. WT3 requires FBC = 0 at EVERY Function TNO in both WT1 and WT2.
- Count remaining BLKs whose current Pattern endpoint FBC > 0. Do not subtract FBC magnitudes or detailed Address counts. A Hard Fail BLK is entirely excluded even if it also contains Soft Fail bits.
- Required missing records yield an unavailable gray/blank Map cell, never a zero/Pass. Parser errors remain yellow/blank. When all 256 BLKs are excluded, the remaining Fail BLK Count is zero.
- Monitor TNOs do not participate. FBC and Fail Address source values are preserved; zero FBC values remain numeric but display blank.

Native Excel validation covers two Hard Fail TNOs, BLK 0/255 boundaries, all three WTs, preceding-WT all-Function filters (including an early non-Stop TNO that fails while the final TNO is zero), permanent exclusion, absent prerequisites, no-Hard-Fail mode, and unchanged source values. Actual company ASC files were not provided for this validation.
