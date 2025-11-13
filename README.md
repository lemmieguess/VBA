# BWS - Bridgewater Studio VBA Macro

**Version:** v1.7.13
**Status:** ✅ Production Ready
**Last Updated:** 2025-11-13

---

## 🚀 Quick Start

**Current Production Version:** `BWS_v1.7.13.bas`

**Development Branch:** `claude/main-cleanup-011CV5RDnCuDivzcYKocKgFF`

### For Developers:
1. Read **[REPOSITORY_STATUS.md](REPOSITORY_STATUS.md)** - Complete repo overview
2. Review **[VBA_CODING_STANDARDS.md](VBA_CODING_STANDARDS.md)** - Required before coding
3. Check **[PRE_COMMIT_CHECKLIST.md](PRE_COMMIT_CHECKLIST.md)** - Before every commit

### For Non-Coders:
- Start with **[README_FOR_NON_CODERS.md](README_FOR_NON_CODERS.md)**
- Use **[PRINTABLE_CHECKLIST.md](PRINTABLE_CHECKLIST.md)** as quick reference

---

## 📋 What is BWS?

BWS (Bridgewater Studio) is a **Microsoft Word VBA macro** that automates document formatting and workflow for business letters, proposals, and reports.

### Key Features:
- ✅ **Import & Format** - Import drafts into branded letterhead templates
- ✅ **Smart Tables** - Auto-format with headers, zebra striping, currency alignment
- ✅ **Bullet Conversion** - Convert text bullets to proper Word formatting
- ✅ **Signature Management** - Handle signature blocks with image positioning
- ✅ **Persistent Toolbar** - 3-row toolbar with all BWS functions
- ✅ **One-Click Workflow** - Import newest draft with single button press

---

## 🆕 Latest Changes (v1.7.13)

### Critical Fixes:
1. **Table Indentation** - Fixed negative right indent (-0.25")
   - Added `RightIndent = 0` for both rows and cells
   - Added table-level `ParagraphFormat.RightIndent = 0`

2. **Cell Spacing** - Ensured zero spacing in all table cells
   - `SpaceBefore = 0` and `SpaceAfter = 0`

3. **Image Preservation** - Signature images no longer deleted
   - Check `InlineShapes.Count > 0` before deleting paragraphs

### See Also:
- Full version history: [REPOSITORY_STATUS.md](REPOSITORY_STATUS.md#-version-history-summary)
- v1.7.12: Fixed Normal.dotm toolbar duplication
- v1.7.11: Fixed Dim placement in ExtractValue function

---

## 📁 Repository Structure

```
VBA/
├── README.md                      ← YOU ARE HERE
├── REPOSITORY_STATUS.md           ← Complete repository overview ⭐
│
├── BWS_v1.7.13.bas               ← CURRENT PRODUCTION CODE ⭐
├── BWS_v1.7.11.bas               ← Previous stable version
├── BWS_v1.7.10.bas               ← Reference version
│
├── VBA_CODING_STANDARDS.md       ← Required reading before coding
├── DEVELOPMENT_WORKFLOW.md       ← Non-coder friendly guide
├── PRE_COMMIT_CHECKLIST.md       ← Check before every commit
├── PRINTABLE_CHECKLIST.md        ← Quick reference card
├── README_FOR_NON_CODERS.md      ← User documentation
│
├── check_vba_code.py             ← Automated validator
├── validate_code.bat             ← Windows validation script
│
└── [Test Documents & Images]
```

---

## 🔧 Installation

### Prerequisites:
- Microsoft Word (Windows or Mac)
- VBA enabled (Developer tab visible)

### Steps:
1. Download `BWS_v1.7.13.bas` from this repository
2. Open Microsoft Word
3. Press `Alt+F11` (Windows) or `Fn+Opt+F11` (Mac) to open VBA Editor
4. File → Import File → Select `BWS_v1.7.13.bas`
5. Close VBA Editor
6. Run `BWS_Install` macro to configure:
   - Base folder (Dropbox root)
   - Drafts folder
   - Letterhead template (.dotm/.dotx)

### Usage:
- A persistent 3-row toolbar appears in Word
- Click "Import Newest" to import and format latest draft
- Click "New Letter" to create blank letterhead
- All formatting is automatic!

---

## 🛠️ Development

### Before Making Changes:

```bash
# Clone the repository
git clone https://github.com/lemmieguess/VBA.git
cd VBA

# Checkout the canonical branch
git checkout claude/main-cleanup-011CV5RDnCuDivzcYKocKgFF

# Review coding standards
cat VBA_CODING_STANDARDS.md
```

### Making Changes:

1. **Edit** `BWS_vX.X.X.bas` in VBA Editor
2. **Update version** in 3 places:
   - `BWS_VERSION` constant (line ~112)
   - `BWS_About()` function (line ~1670)
   - Your commit message
3. **Validate** before committing:
   ```bash
   python3 check_vba_code.py BWS_vX.X.X.bas
   ```
4. **Test in Word** - Actually run the macro!

### Committing:

```bash
git add .
git commit -m "vX.X.X - Description of changes"
git push origin your-branch-name
```

**See [DEVELOPMENT_WORKFLOW.md](DEVELOPMENT_WORKFLOW.md) for detailed guide**

---

## 📚 Documentation Guide

| File | Purpose | Audience |
|------|---------|----------|
| **README.md** | Project overview | Everyone |
| **REPOSITORY_STATUS.md** | Complete repo state | Developers ⭐ |
| **VBA_CODING_STANDARDS.md** | Critical coding rules | Developers (MUST READ) |
| **DEVELOPMENT_WORKFLOW.md** | Step-by-step guide | Non-coders |
| **PRE_COMMIT_CHECKLIST.md** | Pre-commit checks | Developers |
| **PRINTABLE_CHECKLIST.md** | Quick reference | Everyone |
| **README_FOR_NON_CODERS.md** | User guide | End users |

**Start here:** [REPOSITORY_STATUS.md](REPOSITORY_STATUS.md)

---

## 🚨 Common Errors & Solutions

### Error: "Compile Error: Expected End of Statement"
- **Cause:** Dim statement inside loop/If block
- **Solution:** Move ALL Dim statements to function level
- **Reference:** VBA_CODING_STANDARDS.md §2

### Error: "Too many line continuations"
- **Cause:** More than 25 `_` continuations
- **Solution:** Split into multiple variables (max 20 continuations each)
- **Reference:** VBA_CODING_STANDARDS.md §1

### Error: "Paragraph.Index property not found"
- **Cause:** Paragraph objects don't have .Index
- **Solution:** Use `Range.MoveEnd` instead
- **Reference:** VBA_CODING_STANDARDS.md §6

### Error: Version numbers inconsistent
- **Cause:** Forgot to update version in all 3 locations
- **Solution:** Update BWS_VERSION, BWS_About(), and commit message
- **Reference:** PRE_COMMIT_CHECKLIST.md

---

## ✅ Validation

All code is automatically validated before commits:

```bash
python3 check_vba_code.py BWS_v1.7.13.bas
```

**Checks:**
- ✅ Line continuations (<20)
- ✅ Dim statement placement (function level only)
- ✅ Version number consistency
- ✅ Common VBA errors

**Current Status:** 🎉 ALL CHECKS PASSED (4/4)

---

## 📊 Version Timeline

| Version | Date | Summary |
|---------|------|---------|
| **v1.7.13** | **2025-11-13** | **Table indent & image preservation** ⭐ |
| v1.7.12 | 2025-10-24 | Normal.dotm toolbar duplication fix |
| v1.7.11 | 2025-10-24 | ExtractValue Dim placement fix |
| v1.7.10 | 2025-10-24 | Signature block improvements |
| v1.7.0-v1.7.9 | 2025-10 | Progressive bug fixes |
| v1.6.6 | Earlier | Bullet & table alignment fixes |
| v1.6.0 | Earlier | Major feature release |

---

## 🤝 Contributing

### Branch Strategy:
- **Main Development:** `claude/main-cleanup-011CV5RDnCuDivzcYKocKgFF`
- **Feature Branches:** `claude/feature-name-[SESSION_ID]`
- **Naming:** Must start with `claude/` and end with session ID

### Workflow:
1. Create branch from canonical branch
2. Make changes following VBA_CODING_STANDARDS.md
3. Run validation: `python3 check_vba_code.py`
4. Test in Word
5. Commit with version update
6. Push and create PR

**See [DEVELOPMENT_WORKFLOW.md](DEVELOPMENT_WORKFLOW.md) for complete workflow**

---

## 📞 Support

**Issues?** Check these resources:
1. [REPOSITORY_STATUS.md](REPOSITORY_STATUS.md) - Repository overview
2. [VBA_CODING_STANDARDS.md](VBA_CODING_STANDARDS.md) - Coding rules
3. [README_FOR_NON_CODERS.md](README_FOR_NON_CODERS.md) - User guide

**Questions?**
- GitHub Issues: https://github.com/lemmieguess/VBA/issues
- See documentation files in repository root

---

## 📜 License

Internal use for Bridgewater Studio.

---

## 🎯 Quick Links

- **[Complete Repository Status](REPOSITORY_STATUS.md)** - Full overview
- **[Coding Standards](VBA_CODING_STANDARDS.md)** - Must read before coding
- **[Pre-Commit Checklist](PRE_COMMIT_CHECKLIST.md)** - Every commit
- **[Development Workflow](DEVELOPMENT_WORKFLOW.md)** - Step-by-step guide
- **[Current Code](BWS_v1.7.13.bas)** - v1.7.13 (production)

---

**Project:** BWS (Bridgewater Studio) VBA Macro Development
**Repository:** https://github.com/lemmieguess/VBA
**Maintained By:** Claude Code
**Last Updated:** 2025-11-13
