# Images used on the site

All eight images below ship with the project and are referenced directly from the CSS. The page will look complete on first load without any further action. If you want to swap any image, keep the filename the same and the CSS will pick it up.

All images are sourced from Pexels, which permits free commercial and editorial use without attribution. URLs below point to the Pexels page, not the raw file, so you can see the license and photographer.

| Filename           | Section                     | Pexels ID   | Source                                                      |
| ------------------ | --------------------------- | ----------- | ----------------------------------------------------------- |
| `hero.jpg`         | Hero / Invocation           | `1666816`   | https://www.pexels.com/photo/1666816/                       |
| `encounter.jpg`    | The song (video)            | `34612562`  | https://www.pexels.com/photo/34612562/                      |
| `context.jpg`      | Context                     | `7520072`   | https://www.pexels.com/photo/choir-singing-in-a-church-7520072/ |
| `congregation.jpg` | What the song does          | `7520361`   | https://www.pexels.com/photo/choir-singing-in-a-church-7520361/ |
| `prayer.jpg`       | What the song says (pivot)  | `14210110`  | https://www.pexels.com/photo/man-praying-with-his-eyes-closed-14210110/ |
| `sanctuary.jpg`    | Practical theology          | `15666878`  | https://www.pexels.com/photo/church-interior-with-stained-glass-windows-15666878/ |
| `procession.jpg`   | Belief in action            | `33763186`  | https://www.pexels.com/photo/humanitarian-aid-distribution-in-african-village-33763186/ |
| `candle.jpg`       | Benediction                 | `35890634`  | https://www.pexels.com/photo/peaceful-church-candles-illuminating-the-dark-35890634/ |

## Why each image was chosen

- **hero.jpg** A wide shot of a worshipping crowd with hands raised. This is the opening invocation. The image establishes communal energy before the thesis is even read.
- **encounter.jpg** A silhouetted worshipper against stage light. Sits behind the YouTube embed with a heavy black overlay so the video stays the focal point while the viewer feels placed *inside* a worship service.
- **context.jpg** A gospel choir mid-song in a sunlit church. Pairs with the Sinach biography to show worship as it actually happens, not as a staged image of a worshipper.
- **congregation.jpg** The same choir in movement. Used as a background behind the `What the song does` numbered list to ground the claim that the song is a thing bodies do together, not just hear.
- **prayer.jpg** Close-up of a man in prayer, eyes closed. Anchors the `Identity claimed in the first person` section. The lyric claims (I am holy, I am righteous) are harder to read abstractly once a face is next to them.
- **sanctuary.jpg** A cathedral interior with stained-glass light. The practical theology section is the contemplative hinge of the argument, and the image literalizes sacred space. It sits beneath the halo and ray treatments at low opacity.
- **procession.jpg** Humanitarian aid being distributed in an African village. This is the only image in the deck that leaves the sanctuary. The `Belief in action` claim gets a visual: worship moves into public life.
- **candle.jpg** Candles burning in low light. Closing image. Dark, warm, intimate, like the moment after the final blessing.

## Style notes for replacements

- The palette assumes amber, clay, oxblood, and gold. Warm, earthy images work best.
- Dim or moody photographs are welcome because the CSS applies its own gradient wash. Overexposed or neon images will fight the layout.
- Avoid images with visible text, watermarks, or very busy foreground subjects. The page sets type on top of several of these images.
- Minimum width 1200 pixels for side panels, 1600 for full-section backgrounds.

## Reproducing the downloads

Pexels exposes the compressed CDN file at a predictable URL. The command below re-downloads any image at a requested width:

```powershell
$id = "1666816"; $w = 1920; $out = "images/hero.jpg"
Invoke-WebRequest -Uri "https://images.pexels.com/photos/$id/pexels-photo-$id.jpeg?auto=compress&cs=tinysrgb&w=$w" -OutFile $out
```

## Licensing

Pexels content is distributed under the [Pexels License](https://www.pexels.com/license/), which permits free use including commercial use and does not require attribution. Attribution is still the polite thing to do, and the table above doubles as a credit list.
