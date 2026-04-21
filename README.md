# I Know Who I Am: Presentation Package

A cohesive presentation package on Sinach's *I Know Who I Am* for the Yale Divinity School course *African Urban Worship Music*. Three deliverables share the same argument, structure, and bibliography so that the slides, the spoken script, and the webpage all tell the same story.

**Central argument**

> *I Know Who I Am* is not just a song about identity. It is a practical theological act that forms believers to live differently in the world. By teaching worshippers who they are before God, the song gives them confidence, dignity, and a sense of responsibility that can move belief into action.

---

## File tree

```
AUWM Final Presentation/
├── I_Know_Who_I_Am_Presentation.pptx   # 9-slide deck with speaker notes
├── I_Know_Who_I_Am_Full_Script.docx    # ~960-word spoken script, 8 minutes
├── index.html                          # Webpage version with embedded video
├── styles.css                          # Webpage styling
├── script.js                           # Scroll animation + active nav
├── generate_assets.py                  # Script that builds the .pptx and .docx
└── README.md                           # This file
```

---

## How to view each deliverable

### 1. PowerPoint deck

Open `I_Know_Who_I_Am_Presentation.pptx` in PowerPoint or Keynote. The deck is 16:9 and includes:

- A title slide
- Seven content slides (introduction, context, features, individual identity, collective identity, why it matters, conclusion)
- A bibliography slide
- Speaker notes on every slide (also available under the *Notes* pane in PowerPoint)

### 2. Word script

Open `I_Know_Who_I_Am_Full_Script.docx` in Word, Google Docs, or Pages. The script:

- Runs about 8 minutes at a natural speaking pace
- Is roughly 990 words
- Moves from identity to practical theology to action
- Uses bracketed slide cues like `[Slide 1]`, `[Slide 2]` for easy navigation
- Ends with a bibliography page

### 3. Webpage

Open `index.html` directly in a browser. For the cleanest experience, serve the folder locally so the YouTube embed and fonts load over HTTPS:

```powershell
# From the project folder
python -m http.server 8000
# then visit http://localhost:8000
```

The page is responsive, has smooth-scroll navigation, and embeds the Sinach video from YouTube (nocookie domain) so it loads without third party cookies in most browsers.

---

## Rebuilding the PowerPoint or Word file

Both Office files are generated from `generate_assets.py`, which keeps the slide copy, the script, and the bibliography in one place. To regenerate:

```powershell
python -m pip install python-pptx python-docx
python generate_assets.py
```

Edit the `SLIDES`, `SCRIPT`, and `BIBLIOGRAPHY` blocks near the top of the file to change shared content, then re-run the script.

---

## Manual polish checklist

The package is ready to present as-is. If you want to tune it further before submitting:

1. **Verify the placeholder bibliography entry.** The final entry on the bibliography slide, webpage, and Word document is labeled `verify before submitting`. Replace it with a confirmed course reading, or delete it if the other three sources are sufficient.
2. **Add your own image (optional).** The deck is intentionally type-driven. If the course expects a visual, drop an image onto slide 1 or slide 3 in PowerPoint. Good candidates: a wide shot of a worship service, a tasteful photo of Sinach performing, or an abstract warm-toned photo of hands raised in a congregation.
3. **Confirm slide pacing.** At 8 minutes over 9 slides, aim for roughly 50 to 60 seconds per content slide. The speaker notes suggest emphasis points if you need to cut on the fly.
4. **Update the presenter name.** The name *Daniel Baglini* is hard-coded in the deck, script, and webpage. Change it in `generate_assets.py` and `index.html` if needed, then rerun `python generate_assets.py`.

---

## Structure (mirrored across all three deliverables)

1. Title
2. Introduction: more than a song, a practical theological act
3. Song and artist context
4. What the song does: music as formation
5. What the song says: identity claimed in the first person
6. Practical theology: from worship to action (names identity, forms disposition, shapes practice)
7. Why it matters: belief moves into action
8. Conclusion: faith formed, faith lived
9. Bibliography

Slide 6 is the hinge of the argument and is visually distinct from the rest of the deck (dark background, gold accents, three numbered cards).

---

## Sources used

1. Ajose, Toyin. "Multitracking the Spirit: Musicality, Prayer Chants and Pentecostal Spirituality in Nigeria." Harvard Divinity School, 2025.
2. Ajose, Toyin Samuel. "Liturgical traditions and transitions: Congregational hymn singing in Nigerian Pentecostal churches." *Journal of the Association of Nigerian Musicologists*, 2025.
3. Ajose, Toyin Samuel. "'Me I no go suffer': Christian songs and prosperity gospel among Yoruba Pentecostals in Southwest Nigeria."
4. Ayodeji, Oluwafemi Emmanuel. *Ecstasy, Holiness and Spiritual Warfare: Yorùbá Pentecostal Music Experience*. Durham University, 2019.
5. Course materials and lecture framing. *African Urban Worship Music*. Yale Divinity School.
6. King, Roberta Rose, Jean Ngoya Kidula, James R. Krabill, and Thomas Oduro. *Music in the Life of the African Church*. Waco, TX: Baylor University Press, 2008.
7. Rotimi, Oti Alaba. "The Role of Music in African Pentecostal Churches in Southwestern Nigeria."
8. Sinach. "I Know Who I Am." Official music video, 2014. <https://www.youtube.com/watch?v=frtZ4XfoXxM>.

Entries are listed alphabetically by author, with the same numbering used on the bibliography slide, the Word document bibliography page, and the `#bibliography` section of the webpage. No page numbers are fabricated; where a specific publication venue was not verified, the entry names only the author and title.
