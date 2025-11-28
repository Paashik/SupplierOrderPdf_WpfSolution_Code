# Controls Showcase Window Design

## Overview
Create a new WPF window named `ControlsShowcaseWindow.xaml` that displays all UI elements mentioned in the UI improvement prompt for visual style verification. The window will use Material Design 3 styling and adaptive layout.

## Window Properties
- Title: "UI Controls Showcase"
- Width: 1200, Height: 800 (initial, resizable)
- MinWidth: 800, MinHeight: 600
- WindowStartupLocation: CenterScreen
- ResizeMode: CanResize
- Icon: Use app icon

## Layout Structure
Use a ScrollViewer containing a StackPanel with Margin="16" for overall padding.

Each section is a MaterialDesign:Card with Header and content.

Sections in order:
1. Filters Section
2. Buttons Section
3. DataGrid Section
4. Preview Section
5. Cards Section
6. Tabs and Forms Section

## Controls to Include

### Filters Section
- TextBox with search icon (md:HintAssist.Hint="Search")
- ComboBox for Year (items: 2020-2025)
- ComboBox for Status (items: Active, Inactive, Pending)
- Button "Apply Filters" (Filled)

Arranged in a WrapPanel for responsive layout.

### Buttons Section
Demonstrate button variants:
- Button "Primary Action" (Style: MaterialDesignFilledButton)
- Button "Secondary Action" (Style: MaterialDesignFilledTonalButton)
- Button "Tertiary Action" (Style: MaterialDesignOutlinedButton)
- Button "Text Action" (Style: MaterialDesignTextButton)

Arranged horizontally with consistent spacing.

### DataGrid Section
- DataGrid with sample data (columns: ID, Name, Status, Date)
- Styled with Material Design: alternating rows, hover effects, selection.

### Preview Section
- Card with title "Preview Pane"
- WebView2 control inside, loaded with a sample HTML or URL.

### Cards Section
- Two Cards: one for Customer, one for Supplier
- Each with title, content, and action buttons.

### Tabs and Forms Section
- TabControl with two tabs: "General", "Advanced"
- In each tab, Grid with labels and controls:
  - TextBox for SMTP Server
  - PasswordBox for Password
  - CheckBox for SSL
  - Helper text using md:HintAssist.HelperText

## Styling
- Use consistent 8dp spacing (Margin="8")
- Align labels and controls
- Ensure responsive: use star sizing where appropriate
- Typography: use Material Design text styles

## Code-Behind
- Basic data binding for DataGrid
- Sample data population
- Event handlers if needed for interactivity