# Prompt for Improving the WPF UI (Material Design 3)

## Observed visual and UX problems
- **Fixed window metrics hurt responsiveness.** `MainWindow` is hard-coded to 1180×790, so it scales poorly on small/large displays and ignores MD3 density guidance. Login dialog similarly fixes width and removes the native chrome without offering window controls beyond a close button.
- **Filters sit in a single horizontal line with cramped spacing.** Search, year, status, and actions are packed into one row; widths are inconsistent (e.g., narrow year combo vs. wide search box) and vertical alignment jumps because of differing heights derived from other controls. This reduces scannability and breaks MD3 layout rhythm.
- **DataGrid lacks MD3 styling.** The grid uses default headers, cell padding, and focus visuals that clash with the MD3 palette and typography applied elsewhere, making it look like a different toolkit.
- **Preview pane has sparse affordances.** The `WebView2` preview starts 16 px lower than the cards, has no header, and the error message simply toggles visibility without a supporting call-to-action, leaving the area visually empty when content loads or fails.
- **Action buttons and cards compete for emphasis.** Mixed button variants (tonal, filled, outlined, text) sit adjacent without hierarchy; cards use generous padding but typography and spacing vary between customer/supplier sections, creating visual noise.
- **Settings tabs feel dense.** Long forms place labels above controls but mix icon hints and plain fields, producing uneven text alignment; elevation and background colors for cards and tab host are similar, so sections visually blend.

## Questions to clarify before redesign
1. Should the main window support user-resizable layout with remembered sizes, or keep a minimum size only?
2. Which filters are mandatory and which can be tucked under a collapsible “Дополнительно” block to reduce horizontal clutter?
3. Do we need consistent header actions (e.g., refresh, settings) repeated per view, or can they move into an app bar?
4. What level of detail should the orders grid show by default (columns, row density), and do rows need inline actions?
5. For the preview: should it show a placeholder skeleton/spinner while loading, and should errors include a “Установить WebView2” link?
6. For login: is a frameless window required, or can we rely on system chrome with accent-colored title bar to improve accessibility?
7. In Settings, can related SMTP/IMAP fields be grouped into visual sections with helper text instead of long single columns?

## Repair prompt (hand this to the UI author or design assistant)
“Rewrite the WPF XAML to align with Material Design 3:
- Use adaptive layout: remove fixed window sizes, set sensible `MinWidth/MinHeight`, enable `SizeToContent` where appropriate, and make grids use star sizing with consistent 8/4 dp spacing.
- Rebuild the filter area into two rows or a `WrapPanel` with uniform control heights, adding icon-leading hints and consistent label alignment.
- Restyle the `DataGrid` (headers, row padding, hover/focus states, selection color) to match MD3 colors and typography; add empty-state text.
- Give the preview pane a card with a title bar, loading placeholder, and actionable error state (install/help link); align its top margin with the cards above.
- Normalize button hierarchy: primary action uses `Filled`, secondary uses `Tonal`, tertiary uses `Outlined` or `Text`; keep consistent widths and spacing.
- Standardize card padding (16–24 dp) and typography levels across customer/supplier and settings sections; add subtle dividers or subtitles for subsections.
- In Settings, group fields into `Grid` rows with equal label column width, keep helper text in `BodySmall`, and reduce visual noise by avoiding redundant icons.
- Restore accessible window chrome unless frameless is required; if frameless, provide drag regions and keyboard-focusable close/minimize buttons.”
