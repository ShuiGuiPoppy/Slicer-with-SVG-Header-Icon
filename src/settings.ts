"use strict";

import powerbi from "powerbi-visuals-api";
import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsModel = formattingSettings.Model;
import FormattingSettingsSlice = formattingSettings.Slice;

const itemAlignmentOptions: powerbi.IEnumMember[] = [
    { value: "left", displayName: "Left" },
    { value: "center", displayName: "Center" },
    { value: "right", displayName: "Right" }
];

class HeaderCardSettings extends FormattingSettingsCard {
    titleText = new formattingSettings.TextInput({
        name: "titleText",
        displayName: "Title text",
        value: "",
        placeholder: "Category"
    });

    showTitle = new formattingSettings.ToggleSwitch({
        name: "showTitle",
        displayName: "Show title",
        value: true
    });

    selectedValueText = new formattingSettings.TextInput({
        name: "selectedValueText",
        displayName: "Value placeholder",
        value: "All",
        placeholder: "All"
    });

    backgroundColor = new formattingSettings.ColorPicker({
        name: "backgroundColor",
        displayName: "Background colour",
        value: undefined,
        isNoFillItemSupported: true
    });

    textColor = new formattingSettings.ColorPicker({
        name: "textColor",
        displayName: "Text colour",
        value: { value: "#2b3c85" }
    });

    titleFontFamily = new formattingSettings.FontPicker({
        name: "titleFontFamily",
        value: "Segoe UI"
    });

    titleFontSize = new formattingSettings.NumUpDown({
        name: "titleFontSize",
        value: 14
    });

    titleBold = new formattingSettings.ToggleSwitch({
        name: "titleBold",
        value: false
    });

    titleItalic = new formattingSettings.ToggleSwitch({
        name: "titleItalic",
        value: false
    });

    titleUnderline = new formattingSettings.ToggleSwitch({
        name: "titleUnderline",
        value: false
    });

    titleFont = new formattingSettings.FontControl({
        name: "titleFont",
        displayName: "Font",
        fontFamily: this.titleFontFamily,
        fontSize: this.titleFontSize,
        bold: this.titleBold,
        italic: this.titleItalic,
        underline: this.titleUnderline
    });

    valueFontFamily = new formattingSettings.FontPicker({
        name: "valueFontFamily",
        value: "Segoe UI"
    });

    valueFontSize = new formattingSettings.NumUpDown({
        name: "valueFontSize",
        value: 14
    });

    valueBold = new formattingSettings.ToggleSwitch({
        name: "valueBold",
        value: false
    });

    valueItalic = new formattingSettings.ToggleSwitch({
        name: "valueItalic",
        value: false
    });

    valueUnderline = new formattingSettings.ToggleSwitch({
        name: "valueUnderline",
        value: false
    });

    valueFont = new formattingSettings.FontControl({
        name: "valueFont",
        displayName: "Value font",
        fontFamily: this.valueFontFamily,
        fontSize: this.valueFontSize,
        bold: this.valueBold,
        italic: this.valueItalic,
        underline: this.valueUnderline
    });

    name = "header";
    displayName = "Header";
    slices: Array<FormattingSettingsSlice> = [
        this.titleText,
        this.showTitle,
        this.selectedValueText,
        this.backgroundColor,
        this.textColor,
        this.titleFont,
        this.valueFont
    ];
}

class IconCardSettings extends FormattingSettingsCard {
    private static readonly defaultSvgMarkup: string = [
        "<svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"__LINE_COLOR__\" stroke-width=\"1.8\"",
        " stroke-linecap=\"round\" stroke-linejoin=\"round\" xmlns=\"http://www.w3.org/2000/svg\">",
        "<path d=\"M6 4.5h12a1.5 1.5 0 0 1 1.5 1.5v12A1.5 1.5 0 0 1 18 19.5H6A1.5 1.5 0 0 1 4.5 18V6A1.5 1.5 0 0 1 6 4.5Z\" fill=\"__FILL_COLOR__\"/>",
        "<path d=\"M8 9.25h8\"/><path d=\"M8 12h8\"/><path d=\"M8 14.75h5\"/></svg>"
    ].join("");

    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "Show icon",
        value: true
    });

    svgMarkup = new formattingSettings.TextArea({
        name: "svgMarkup",
        displayName: "SVG markup",
        value: IconCardSettings.defaultSvgMarkup,
        placeholder: "<svg>...</svg>"
    });

    lineColor = new formattingSettings.ColorPicker({
        name: "lineColor",
        displayName: "Line colour",
        value: { value: "#2b3c85" },
        isNoFillItemSupported: true
    });

    fillColor = new formattingSettings.ColorPicker({
        name: "fillColor",
        displayName: "Fill colour",
        value: { value: "#ffffff" },
        isNoFillItemSupported: true
    });

    size = new formattingSettings.NumUpDown({
        name: "size",
        displayName: "Icon size",
        value: 30
    });

    name = "icon";
    displayName = "Icon";
    slices: Array<FormattingSettingsSlice> = [
        this.show,
        this.svgMarkup,
        this.lineColor,
        this.fillColor,
        this.size
    ];
}

class BorderLayoutCardSettings extends FormattingSettingsCard {
    cardBackgroundColor = new formattingSettings.ColorPicker({
        name: "cardBackgroundColor",
        displayName: "Card background colour",
        value: { value: "transparent" }
    });

    borderColor = new formattingSettings.ColorPicker({
        name: "borderColor",
        displayName: "Border colour",
        value: { value: "#2b3c85" }
    });

    borderRadius = new formattingSettings.NumUpDown({
        name: "borderRadius",
        displayName: "Border radius",
        value: 12
    });

    padding = new formattingSettings.NumUpDown({
        name: "padding",
        displayName: "Padding",
        value: 0
    });

    name = "borderLayout";
    displayName = "Border & Layout";
    slices: Array<FormattingSettingsSlice> = [
        this.cardBackgroundColor,
        this.borderColor,
        this.borderRadius,
        this.padding
    ];
}

class ItemsCardSettings extends FormattingSettingsCard {
    itemFontFamily = new formattingSettings.TextInput({
        name: "itemFontFamily",
        displayName: "Font family",
        value: "Segoe UI",
        placeholder: "Segoe UI"
    });

    itemFontSize = new formattingSettings.NumUpDown({
        name: "itemFontSize",
        displayName: "Font size",
        value: 14
    });

    itemAlignment = new formattingSettings.ItemDropdown({
        name: "itemAlignment",
        displayName: "Values alignment",
        items: itemAlignmentOptions,
        value: itemAlignmentOptions[1]
    });

    itemTextColor = new formattingSettings.ColorPicker({
        name: "itemTextColor",
        displayName: "Text colour",
        value: { value: "#2b3c85" }
    });

    backgroundColor = new formattingSettings.ColorPicker({
        name: "backgroundColor",
        displayName: "Background colour",
        value: { value: "#ffffff" }
    });

    rowBackgroundColor = new formattingSettings.ColorPicker({
        name: "rowBackgroundColor",
        displayName: "Row background colour",
        value: { value: "#ffffff" }
    });

    hoverColor = new formattingSettings.ColorPicker({
        name: "hoverColor",
        displayName: "Hover background colour",
        value: { value: "#f3f4f6" }
    });

    selectedBackgroundColor = new formattingSettings.ColorPicker({
        name: "selectedBackgroundColor",
        displayName: "Selected background colour",
        value: { value: "#e1e7fd" }
    });

    dropdownMaxHeight = new formattingSettings.NumUpDown({
        name: "dropdownMaxHeight",
        displayName: "Dropdown max height",
        value: 220
    });

    showSearch = new formattingSettings.ToggleSwitch({
        name: "showSearch",
        displayName: "Show search",
        value: true
    });

    alwaysShowOptions = new formattingSettings.ToggleSwitch({
        name: "alwaysShowOptions",
        displayName: "Always show options",
        value: false
    });

    fitToVisualHeight = new formattingSettings.ToggleSwitch({
        name: "fitToVisualHeight",
        displayName: "Fit to visual height",
        value: false
    });

    name = "items";
    displayName = "Items";
    slices: Array<FormattingSettingsSlice> = [
        this.itemFontFamily,
        this.itemFontSize,
        this.itemAlignment,
        this.itemTextColor,
        this.backgroundColor,
        this.rowBackgroundColor,
        this.hoverColor,
        this.selectedBackgroundColor,
        this.dropdownMaxHeight,
        this.showSearch,
        this.alwaysShowOptions,
        this.fitToVisualHeight
    ];
}

class SelectionCardSettings extends FormattingSettingsCard {
    allowMultiSelect = new formattingSettings.ToggleSwitch({
        name: "allowMultiSelect",
        displayName: "Allow multi-select",
        value: false
    });

    name = "selection";
    displayName = "Selection";
    slices: Array<FormattingSettingsSlice> = [this.allowMultiSelect];
}

class SortCardSettings extends FormattingSettingsCard {
    sortAscending = new formattingSettings.ToggleSwitch({
        name: "sortAscending",
        displayName: "Sort ascending",
        value: true
    });

    name = "sort";
    displayName = "Sort";
    slices: Array<FormattingSettingsSlice> = [this.sortAscending];
}

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    headerCard = new HeaderCardSettings();
    iconCard = new IconCardSettings();
    borderLayoutCard = new BorderLayoutCardSettings();
    itemsCard = new ItemsCardSettings();
    selectionCard = new SelectionCardSettings();
    sortCard = new SortCardSettings();

    cards = [this.headerCard, this.iconCard, this.borderLayoutCard, this.itemsCard, this.selectionCard, this.sortCard];
}
