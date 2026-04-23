"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import "./../style/visual.less";

import DataView = powerbi.DataView;
import DataViewCategoryColumn = powerbi.DataViewCategoryColumn;
import ISelectionId = powerbi.visuals.ISelectionId;
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import PrimitiveValue = powerbi.PrimitiveValue;
import IVisual = powerbi.extensibility.visual.IVisual;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;

import { VisualFormattingSettingsModel } from "./settings";

type Nullable<T> = T | undefined | null;
type SlicerItem = {
    index: number;
    label: string;
    selectionId: ISelectionId;
};

export class Visual implements IVisual {
    private readonly target: HTMLElement;
    private readonly host: IVisualHost;
    private readonly selectionManager: ISelectionManager;
    private readonly formattingSettingsService: FormattingSettingsService;
    private readonly root: HTMLDivElement;
    private readonly header: HTMLDivElement;
    private readonly titleRow: HTMLDivElement;
    private readonly iconWrapper: HTMLDivElement;
    private readonly titleElement: HTMLSpanElement;
    private readonly valueField: HTMLDivElement;
    private readonly valueElement: HTMLSpanElement;
    private readonly chevronElement: HTMLSpanElement;
    private readonly resetButton: HTMLButtonElement;
    private readonly listContainer: HTMLDivElement;
    private readonly searchInput: HTMLInputElement;
    private readonly itemsContainer: HTMLDivElement;
    private readonly emptyStateElement: HTMLDivElement;

    private isDropdownOpen: boolean;
    private searchText: string;
    private selectedKeys: string[];
    private formattingSettings: VisualFormattingSettingsModel;
    private lastSlicerItems: SlicerItem[];
    private lastPlaceholder: string;

    constructor(options: VisualConstructorOptions) {
        this.target = options.element;
        this.host = options.host;
        this.selectionManager = this.host.createSelectionManager();
        this.formattingSettingsService = new FormattingSettingsService();
        this.isDropdownOpen = false;
        this.searchText = "";
        this.selectedKeys = [];
        this.lastSlicerItems = [];
        this.lastPlaceholder = "All";

        this.root = document.createElement("div");
        this.root.className = "icon-slicer";

        this.header = document.createElement("div");
        this.header.className = "icon-slicer__header";

        this.titleRow = document.createElement("div");
        this.titleRow.className = "icon-slicer__title-row";

        this.iconWrapper = document.createElement("div");
        this.iconWrapper.className = "icon-slicer__icon";

        this.titleElement = document.createElement("span");
        this.titleElement.className = "icon-slicer__title";

        this.valueField = document.createElement("div");
        this.valueField.className = "icon-slicer__field";

        this.valueElement = document.createElement("span");
        this.valueElement.className = "icon-slicer__value";

        this.chevronElement = document.createElement("span");
        this.chevronElement.className = "icon-slicer__chevron";
        this.chevronElement.textContent = "\u25be";

        this.resetButton = document.createElement("button");
        this.resetButton.type = "button";
        this.resetButton.className = "icon-slicer__reset";
        this.resetButton.setAttribute("aria-label", "Clear selection");
        this.resetButton.textContent = "\u2715";

        this.listContainer = document.createElement("div");
        this.listContainer.className = "icon-slicer__list";

        this.searchInput = document.createElement("input");
        this.searchInput.className = "icon-slicer__search";
        this.searchInput.type = "search";
        this.searchInput.placeholder = "Search";
        this.searchInput.setAttribute("aria-label", "Search slicer values");

        this.itemsContainer = document.createElement("div");
        this.itemsContainer.className = "icon-slicer__items";

        this.emptyStateElement = document.createElement("div");
        this.emptyStateElement.className = "icon-slicer__empty";

        this.valueField.tabIndex = 0;
        this.valueField.setAttribute("role", "button");

        this.titleRow.append(this.iconWrapper, this.titleElement);
        this.valueField.append(this.valueElement, this.chevronElement, this.resetButton);
        this.header.append(this.titleRow, this.valueField);
        this.listContainer.append(this.searchInput, this.itemsContainer);
        this.root.append(this.header, this.listContainer, this.emptyStateElement);
        this.target.appendChild(this.root);

        this.valueField.addEventListener("click", () => {
            this.isDropdownOpen = !this.isDropdownOpen;
            this.syncDropdownState();
        });

        this.valueField.addEventListener("keydown", (event: KeyboardEvent) => {
            if (event.key === "Enter" || event.key === " ") {
                event.preventDefault();
                this.isDropdownOpen = !this.isDropdownOpen;
                this.syncDropdownState();
            }
        });

        this.resetButton.addEventListener("click", (event: MouseEvent) => {
            event.stopPropagation();
            void this.clearSelection();
        });

        this.searchInput.addEventListener("input", () => {
            this.searchText = this.searchInput.value;
            this.renderList(this.lastSlicerItems, this.selectedKeys);
            this.syncDropdownState();
        });

        this.selectionManager.registerOnSelectCallback((ids: powerbi.extensibility.ISelectionId[]) => {
            this.selectedKeys = this.toSelectionKeys(ids as unknown as ISelectionId[]);
            this.refreshSelectionUi();
        });
    }

    public update(options: VisualUpdateOptions): void {
        const dataView: Nullable<DataView> = options.dataViews?.[0];
        this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(
            VisualFormattingSettingsModel,
            dataView
        );

        const slicerItems = this.getCategoryItems(dataView);
        const selectedIds = (this.selectionManager.getSelectionIds() || []) as unknown as ISelectionId[];
        const headerSettings = this.formattingSettings.headerCard;
        const iconSettings = this.formattingSettings.iconCard;
        const borderSettings = this.formattingSettings.borderLayoutCard;
        const itemsSettings = this.formattingSettings.itemsCard;
        const selectionSettings = this.formattingSettings.selectionCard;
        const sortSettings = this.formattingSettings.sortCard;
        const alwaysShowOptions = itemsSettings.alwaysShowOptions.value;
        const fitToVisualHeight = itemsSettings.fitToVisualHeight.value;
        const showSearch = itemsSettings.showSearch.value;
        const allowMultiSelect = selectionSettings.allowMultiSelect.value;
        const sortAscending = sortSettings.sortAscending.value;
        const hasCategory = slicerItems.length > 0;

        const showTitle = headerSettings.showTitle.value;
        const resolvedTitle = hasCategory
            ? (headerSettings.titleText.value?.trim() || this.getCategoryDisplayName(dataView) || "Category")
            : "";
        const resolvedPlaceholder = headerSettings.selectedValueText.value?.trim() || "All";
        const managerSelectedKeys = this.toSelectionKeys(selectedIds);

        if (!hasCategory) {
            this.selectedKeys = [];
            this.searchText = "";
            this.isDropdownOpen = false;
            if (selectedIds.length > 0) {
                void this.selectionManager.clear();
            }
        } else {
            this.selectedKeys = managerSelectedKeys.length > 0 ? managerSelectedKeys : this.selectedKeys;
        }

        this.lastSlicerItems = slicerItems;
        this.lastPlaceholder = resolvedPlaceholder;
        const selectedLabel = this.getSelectedLabel(slicerItems, this.selectedKeys, resolvedPlaceholder);
        const hasSelection = this.selectedKeys.length > 0;
        const showTitleRow = iconSettings.show.value || showTitle;
        this.titleElement.textContent = resolvedTitle;
        this.titleElement.hidden = !hasCategory || !showTitle;
        this.titleRow.style.display = hasCategory && showTitleRow ? "flex" : "none";
        this.valueElement.textContent = hasCategory ? selectedLabel : "";
        this.emptyStateElement.textContent = hasCategory
            ? ""
            : "Add a field to the Category bucket to populate the slicer.";

        this.root.style.setProperty("--icon-slicer-header-bg", this.getColorValue(headerSettings.backgroundColor.value, "transparent"));
        this.root.style.setProperty("--icon-slicer-card-bg", this.getColorValue(borderSettings.cardBackgroundColor.value, "transparent"));
        this.root.style.setProperty("--icon-slicer-text", this.getColorValue(headerSettings.textColor.value, "#2b3c85"));
        this.root.style.setProperty("--icon-slicer-border", this.getColorValue(borderSettings.borderColor.value, "#2b3c85"));
        this.root.style.setProperty("--icon-slicer-radius", `${this.getNumericValue(borderSettings.borderRadius.value, 12)}px`);
        this.root.style.setProperty("--icon-slicer-padding", `${this.getNumericValue(borderSettings.padding.value, 12)}px`);
        this.root.style.setProperty("--icon-slicer-title-font-family", this.getFontFamilyValue(headerSettings.titleFontFamily.value, "Segoe UI"));
        this.root.style.setProperty("--icon-slicer-title-font-size", `${this.getNumericValue(headerSettings.titleFontSize.value, 14)}px`);
        this.root.style.setProperty("--icon-slicer-title-font-weight", headerSettings.titleBold.value ? "700" : "400");
        this.root.style.setProperty("--icon-slicer-title-font-style", headerSettings.titleItalic.value ? "italic" : "normal");
        this.root.style.setProperty("--icon-slicer-title-text-decoration", headerSettings.titleUnderline.value ? "underline" : "none");
        this.root.style.setProperty("--icon-slicer-value-font-family", this.getFontFamilyValue(headerSettings.valueFontFamily.value, "Segoe UI"));
        this.root.style.setProperty("--icon-slicer-value-font-size", `${this.getNumericValue(headerSettings.valueFontSize.value, 14)}px`);
        this.root.style.setProperty("--icon-slicer-value-font-weight", headerSettings.valueBold.value ? "700" : "400");
        this.root.style.setProperty("--icon-slicer-value-font-style", headerSettings.valueItalic.value ? "italic" : "normal");
        this.root.style.setProperty("--icon-slicer-value-text-decoration", headerSettings.valueUnderline.value ? "underline" : "none");
        this.root.style.setProperty("--icon-slicer-item-font-family", this.getFontFamilyValue(itemsSettings.itemFontFamily.value, "Segoe UI"));
        this.root.style.setProperty("--icon-slicer-item-font-size", `${this.getNumericValue(itemsSettings.itemFontSize.value, 14)}px`);
        this.root.style.setProperty("--icon-slicer-item-align", this.getTextAlignValue(itemsSettings.itemAlignment.value));
        this.root.style.setProperty("--icon-slicer-item-text", this.getColorValue(itemsSettings.itemTextColor.value, "#2b3c85"));
        this.root.style.setProperty("--icon-slicer-dropdown-bg", this.getColorValue(itemsSettings.backgroundColor.value, "#ffffff"));
        this.root.style.setProperty("--icon-slicer-item-bg", this.getColorValue(itemsSettings.rowBackgroundColor.value, "#ffffff"));
        this.root.style.setProperty("--icon-slicer-item-hover", this.getColorValue(itemsSettings.hoverColor.value, "#f3f4f6"));
        this.root.style.setProperty("--icon-slicer-item-selected", this.getColorValue(itemsSettings.selectedBackgroundColor.value, "#cbcbcb"));
        this.root.style.setProperty("--icon-slicer-list-max-height", `${this.getNumericValue(itemsSettings.dropdownMaxHeight.value, 220)}px`);

        const iconLineColor = this.getOptionalColorValue(iconSettings.lineColor.value);
        const iconFillColor = this.getOptionalColorValue(iconSettings.fillColor.value);

        const iconSize = this.getNumericValue(iconSettings.size.value, 18);
        this.root.style.setProperty("--icon-slicer-icon-size", `${iconSize}px`);
        this.iconWrapper.style.width = `${iconSize}px`;
        this.iconWrapper.style.height = `${iconSize}px`;
        this.iconWrapper.style.display = iconSettings.show.value ? "grid" : "none";
        this.replaceIconSvg(null);

        if (iconSettings.show.value) {
            this.replaceIconSvg(this.resolveSvgElement(
                iconSettings.svgMarkup.value,
                iconLineColor || "none",
                iconFillColor || "none"
            ));
        }

        this.searchInput.style.display = showSearch && hasCategory ? "block" : "none";
        this.searchInput.value = this.searchText;
        this.renderList(slicerItems, this.selectedKeys);
        this.emptyStateElement.style.display = hasCategory ? "none" : "block";
        this.header.style.display = hasCategory ? "flex" : "none";
        this.listContainer.style.display = hasCategory ? this.listContainer.style.display : "none";
        this.valueField.style.cursor = hasCategory && !alwaysShowOptions ? "pointer" : "default";
        this.resetButton.style.display = hasSelection ? "grid" : "none";
        this.resetButton.disabled = !hasSelection;
        this.root.classList.toggle("is-list-mode", alwaysShowOptions);
        this.root.classList.toggle("is-fit-height", fitToVisualHeight);
        this.root.classList.toggle("is-multi-select", allowMultiSelect);
        this.root.classList.toggle("is-sort-desc", !sortAscending);
        this.root.classList.toggle("is-empty", !hasCategory);
        this.valueField.setAttribute("aria-disabled", alwaysShowOptions ? "true" : "false");

        if (!hasCategory) {
            this.isDropdownOpen = false;
        }

        this.syncDropdownState();
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }

    private getCategoryItems(dataView: Nullable<DataView>): SlicerItem[] {
        const categoryColumn: Nullable<DataViewCategoryColumn> = dataView?.categorical?.categories?.[0];
        const rawValues = categoryColumn?.values;

        if (!categoryColumn || !rawValues || !rawValues.length) {
            return [];
        }

        return rawValues
            .map((value: PrimitiveValue, index: number) => ({
                index,
                label: this.getCategoryLabel(value),
                selectionId: this.host.createSelectionIdBuilder().withCategory(categoryColumn, index).createSelectionId()
            }));
    }

    private getCategoryDisplayName(dataView: Nullable<DataView>): string {
        return dataView?.categorical?.categories?.[0]?.source?.displayName?.trim() || "";
    }

    private getColorValue(color: Nullable<powerbi.ThemeColorData>, fallback: string): string {
        return color?.value || fallback;
    }

    private getOptionalColorValue(color: Nullable<powerbi.ThemeColorData>): string | null {
        const value = color?.value?.trim();

        return value ? value : null;
    }

    private getNumericValue(value: Nullable<number>, fallback: number): number {
        return typeof value === "number" && Number.isFinite(value) ? value : fallback;
    }

    private getFontFamilyValue(value: Nullable<string>, fallback: string): string {
        return value?.trim() || fallback;
    }

    private getTextAlignValue(value: Nullable<string | powerbi.IEnumMember>): string {
        const rawValue = typeof value === "string" ? value : value?.value;
        const normalized = `${rawValue ?? ""}`.trim().toLowerCase();

        return normalized === "center" || normalized === "right" ? normalized : "left";
    }

    private replaceIconSvg(svgElement: Nullable<SVGElement>): void {
        this.iconWrapper.replaceChildren();

        if (svgElement) {
            this.iconWrapper.appendChild(svgElement);
        }
    }

    private resolveSvgElement(svgMarkup: Nullable<string>, lineColor: string, fillColor: string): SVGElement | null {
        const fallback = [
            `<svg viewBox="0 0 24 24" fill="none" stroke="${lineColor}" stroke-width="1.8"`,
            " stroke-linecap=\"round\" stroke-linejoin=\"round\" xmlns=\"http://www.w3.org/2000/svg\">",
            `<path d="M6 4.5h12a1.5 1.5 0 0 1 1.5 1.5v12A1.5 1.5 0 0 1 18 19.5H6A1.5 1.5 0 0 1 4.5 18V6A1.5 1.5 0 0 1 6 4.5Z" fill="${fillColor}"/>`,
            "<path d=\"M8 9.25h8\"/><path d=\"M8 12h8\"/><path d=\"M8 14.75h5\"/></svg>"
        ].join("");
        const template = (svgMarkup?.trim() || fallback)
            .replace(/__LINE_COLOR__/g, lineColor)
            .replace(/__FILL_COLOR__/g, fillColor)
            .replace(/currentColor/g, lineColor);

        const parsed = new DOMParser().parseFromString(template, "image/svg+xml");
        const svgElement = parsed.documentElement;

        if (svgElement?.tagName?.toLowerCase() !== "svg") {
            return null;
        }

        this.applySvgColors(svgElement as unknown as SVGElement, lineColor, fillColor);

        return svgElement as unknown as SVGElement;
    }

    private applySvgColors(svgElement: SVGElement, lineColor: string, fillColor: string): void {
        const elements = [svgElement, ...Array.from(svgElement.querySelectorAll("*"))];
        const hasStrokeAttributes = elements.some((element: Element) => {
            const stroke = element.getAttribute("stroke");
            return Boolean(stroke && stroke.toLowerCase() !== "none");
        });

        elements.forEach((element: Element) => {
            const currentFill = element.getAttribute("fill");
            const currentStroke = element.getAttribute("stroke");

            if (currentFill && currentFill.toLowerCase() !== "none") {
                element.setAttribute("fill", hasStrokeAttributes ? fillColor : lineColor);
            }

            if (currentStroke && currentStroke.toLowerCase() !== "none") {
                element.setAttribute("stroke", lineColor);
            }
        });
    }

    private renderList(items: SlicerItem[], selectedKeys: string[]): void {
        this.itemsContainer.replaceChildren();

        if (!items.length) {
            return;
        }

        const filteredItems = this.getFilteredItems(items);

        const allButton = document.createElement("button");
        allButton.type = "button";
        allButton.className = "icon-slicer__item";
        if (!selectedKeys.length) {
            allButton.classList.add("is-selected");
        }
        allButton.textContent = "All";
        allButton.addEventListener("click", (event: MouseEvent) => {
            event.stopPropagation();
            void this.clearSelection();
        });
        this.itemsContainer.appendChild(allButton);

        filteredItems.forEach((item: SlicerItem) => {
            const button = document.createElement("button");
            button.type = "button";
            button.className = "icon-slicer__item";
            button.textContent = item.label;

            if (this.isSelected(item.selectionId, selectedKeys)) {
                button.classList.add("is-selected");
            }

            button.addEventListener("click", (event: MouseEvent) => {
                event.stopPropagation();
                void this.selectItem(item.selectionId);
            });

            this.itemsContainer.appendChild(button);
        });
    }

    private syncDropdownState(): void {
        const alwaysShowOptions = this.formattingSettings?.itemsCard?.alwaysShowOptions?.value;
        const hasVisibleItems = this.itemsContainer.childElementCount > 0 || this.searchInput.style.display !== "none";
        const isOpen = (alwaysShowOptions || this.isDropdownOpen) && hasVisibleItems;

        this.root.classList.toggle("is-open", isOpen);
        this.valueField.setAttribute("aria-expanded", isOpen ? "true" : "false");
        this.listContainer.style.display = isOpen ? "flex" : "none";
    }

    private isSelected(selectionId: ISelectionId, selectedKeys: string[]): boolean {
        return selectedKeys.includes(selectionId.getKey());
    }

    private getSelectedLabel(items: SlicerItem[], selectedKeys: string[], placeholder: string): string {
        if (!items.length) {
            return "No data";
        }

        const selectedItems = items.filter((item: SlicerItem) => this.isSelected(item.selectionId, selectedKeys));

        if (!selectedItems.length) {
            return placeholder;
        }

        if (selectedItems.length === 1) {
            return selectedItems[0].label;
        }

        return `${selectedItems.length} selected`;
    }

    private async selectItem(selectionId: ISelectionId): Promise<void> {
        const allowMultiSelect = this.formattingSettings.selectionCard.allowMultiSelect.value;
        const selectionKey = selectionId.getKey();

        if (allowMultiSelect) {
            this.selectedKeys = this.selectedKeys.includes(selectionKey)
                ? this.selectedKeys.filter((key: string) => key !== selectionKey)
                : [...this.selectedKeys, selectionKey];
        } else {
            this.selectedKeys = [selectionKey];
        }

        const nextSelectionIds = this.lastSlicerItems
            .filter((item: SlicerItem) => this.selectedKeys.includes(item.selectionId.getKey()))
            .map((item: SlicerItem) => item.selectionId);

        if (nextSelectionIds.length === 0) {
            await this.selectionManager.clear();
        } else {
            await this.selectionManager.select(nextSelectionIds, false);
        }

        this.refreshSelectionUi();

        if (!allowMultiSelect) {
            this.isDropdownOpen = false;
        }

        this.syncDropdownState();
    }

    private async clearSelection(): Promise<void> {
        this.selectedKeys = [];
        await this.selectionManager.clear();
        this.refreshSelectionUi();
        this.isDropdownOpen = false;
        this.syncDropdownState();
    }

    private toSelectionKeys(selectionIds: ISelectionId[]): string[] {
        return selectionIds
            .map((selectionId: ISelectionId) => selectionId?.getKey?.())
            .filter((key: string | undefined): key is string => Boolean(key));
    }

    private getFilteredItems(items: SlicerItem[]): SlicerItem[] {
        const normalizedSearch = this.searchText.trim().toLowerCase();
        const orderedItems = this.getItemsWithSelectedFirst(this.getSortedItems(items));

        if (!normalizedSearch) {
            return orderedItems;
        }

        return orderedItems.filter((item: SlicerItem) => item.label.toLowerCase().includes(normalizedSearch));
    }

    private getCategoryLabel(value: PrimitiveValue): string {
        const normalized = `${value ?? ""}`.trim();

        return normalized.length > 0 ? normalized : "(Blank)";
    }

    private getItemsWithSelectedFirst(items: SlicerItem[]): SlicerItem[] {
        const selectedItems: SlicerItem[] = [];
        const unselectedItems: SlicerItem[] = [];

        items.forEach((item: SlicerItem) => {
            if (this.selectedKeys.includes(item.selectionId.getKey())) {
                selectedItems.push(item);
            } else {
                unselectedItems.push(item);
            }
        });

        return [...selectedItems, ...unselectedItems];
    }

    private getSortedItems(items: SlicerItem[]): SlicerItem[] {
        const sortAscending = this.formattingSettings.sortCard.sortAscending.value;
        const direction = sortAscending ? 1 : -1;

        return [...items].sort((left: SlicerItem, right: SlicerItem) =>
            left.label.localeCompare(right.label, undefined, { numeric: true, sensitivity: "base" }) * direction
        );
    }

    private refreshSelectionUi(): void {
        this.valueElement.textContent = this.getSelectedLabel(this.lastSlicerItems, this.selectedKeys, this.lastPlaceholder);
        this.renderList(this.lastSlicerItems, this.selectedKeys);
        this.resetButton.style.display = this.selectedKeys.length > 0 ? "grid" : "none";
        this.resetButton.disabled = this.selectedKeys.length === 0;
        this.syncDropdownState();
    }
}
