/*
*  Power BI Visual CLI
*
*  Copyright (c) Microsoft Corporation
*  All rights reserved.
*  MIT License
*
*  Permission is hereby granted, free of charge, to any person obtaining a copy
*  of this software and associated documentation files (the ""Software""), to deal
*  in the Software without restriction, including without limitation the rights
*  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
*  copies of the Software, and to permit persons to whom the Software is
*  furnished to do so, subject to the following conditions:
*
*  The above copyright notice and this permission notice shall be included in
*  all copies or substantial portions of the Software.
*
*  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
*  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
*  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
*  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
*  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
*  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
*  THE SOFTWARE.
*/
"use strict";

import "regenerator-runtime/runtime";
import "core-js/stable";
import "./../style/visual.less";
import powerbi from "powerbi-visuals-api";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import EnumerateVisualObjectInstancesOptions = powerbi.EnumerateVisualObjectInstancesOptions;
import VisualObjectInstance = powerbi.VisualObjectInstance;
import DataView = powerbi.DataView;
import VisualObjectInstanceEnumerationObject = powerbi.VisualObjectInstanceEnumerationObject;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import DataViewObjects = powerbi.DataViewObjects;
import IVisualEventService = powerbi.extensibility.IVisualEventService;
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import ISelectionId = powerbi.visuals.ISelectionId;
import * as sanitizeHtml from "sanitize-html";
import * as d3 from 'd3';
import { VisualSettings } from "./settings";
import * as validDataUrl from "valid-data-url";

export interface TimelineData {
    Company: String;
    Type: string;
    Description: string;
    Date: Date;
    DocumentLink: string;
    HeaderImage: string;
    FooterImage: string;
    selectionId: powerbi.visuals.ISelectionId;
}

export interface Timelines {
    Timeline: TimelineData[];
}

export class Visual implements IVisual {
    private target: d3.Selection<HTMLElement, any, any, any>;
    private header: d3.Selection<HTMLElement, any, any, any>;
    private footer: d3.Selection<HTMLElement, any, any, any>;
    private svg: d3.Selection<SVGElement, any, any, any>;
    private tooltip: d3.Selection<HTMLElement, any, any, any>;
    private margin = { top: 50, right: 40, bottom: 50, left: 40 };
    private settings: VisualSettings;
    private host: IVisualHost;
    private initLoad = false;
    private events: IVisualEventService;
    private xScale: d3.ScaleTime<number, number>;
    private yScale: d3.ScaleLinear<number, number>;
    private selectionManager: ISelectionManager;

    constructor(options: VisualConstructorOptions) {
        this.target = d3.select(options.element);
        this.header = d3.select(options.element).append("div");
        this.footer = d3.select(options.element).append("div");
        this.svg = d3.select(options.element).append('svg');
        this.tooltip = d3.select(options.element).append("div");
        this.host = options.host;
        this.events = options.host.eventService;
        this.selectionManager = options.host.createSelectionManager();
    }

    public update(options: VisualUpdateOptions) {
        this.events.renderingStarted(options);
        this.settings = Visual.parseSettings(options && options.dataViews && options.dataViews[0]);
        this.svg.selectAll('*').remove();
        this.header.selectAll("img").remove();
        this.header.classed("header", false);
        this.footer.selectAll("img").remove();
        this.footer.classed("footer", false);

        let vpWidth = options.viewport.width;
        let vpHeight = options.viewport.height;

        if (this.settings.dataPoint.layout.toLowerCase() === "header" || this.settings.dataPoint.layout.toLowerCase() === "footer") {
            vpHeight = options.viewport.height - 105;
        }

        let _this = this;
        this.svg.attr('height', vpHeight);
        this.svg.attr('width', vpWidth);

        let gHeight = vpHeight - this.margin.top - this.margin.bottom;
        let gWidth = vpWidth - this.margin.left - this.margin.right;

        this.target.on("contextmenu", () => {
            const mouseEvent: MouseEvent = <MouseEvent> d3.event;
            const eventTarget: any = mouseEvent.target;
            let dataPoint: any = d3.select(eventTarget).datum();
            this.selectionManager.showContextMenu(
              dataPoint ? dataPoint.selectionId : {},
              {
                x: mouseEvent.clientX,
                y: mouseEvent.clientY,
              }
            );
            mouseEvent.preventDefault();
          });

        let timelineData = Visual.CONVERTER(options.dataViews[0], this.host);
        timelineData = timelineData.slice(0, 100);

        let minDate, maxDate, currentDate;
        let timelineLocalData: TimelineData[] = [];
        currentDate = new Date();
        if (timelineData.length > 0) {
            minDate = new Date(currentDate.getFullYear() - this.settings.displayYears.PreviousYear, 0, 1);
            timelineLocalData = timelineData.map<TimelineData>((d) => { if (d.Date.getFullYear() >= minDate.getFullYear()) { return d;} }).filter(e => e);
            maxDate = new Date(currentDate.getFullYear() + this.settings.displayYears.FutureYear, 0, 1);
            timelineLocalData = timelineLocalData.map<TimelineData>((d) => { if (d.Date.getFullYear() <= maxDate.getFullYear()) { return d; } }).filter(e => e);
        }

        if (timelineLocalData.length > 0) {
            timelineData = timelineLocalData;
          } else if (timelineLocalData.length == 0) {
            minDate = new Date(Math.min.apply(null, timelineData.map(d => d.Date)));
            maxDate = new Date(Math.max.apply(null, timelineData.map(d => d.Date)));
            minDate = new Date(minDate.getFullYear(), 0, 1);
            maxDate = new Date(maxDate.getFullYear() + 1, 0, 1);
        }

        this.renderHeaderAndFooter(timelineData);
        this.renderXandYAxis(minDate, maxDate, gWidth, gHeight);
        this.renderTitle(vpWidth, gWidth);
        this.renderLine(timelineData, gHeight);
        this.renderArrow(timelineData);
        this.renderBox(timelineData, gWidth, gHeight);
        this.svg.append('rect')
            .attr('class', 'border-rect')
            .attr('x', 0)
            .attr('y', 0)
            .attr('width', vpWidth)
            .attr('height', vpHeight + 10)
            .attr('stroke-width', '2px')
            .attr('stroke', '#333')
            .attr('fill', 'transparent');
        this.events.renderingFinished(options);
    }

    private renderHeaderAndFooter(timelineData: TimelineData[]) {
        let [timeline] = timelineData;
        if (this.settings.dataPoint.layout.toLowerCase() === "header") {
            this.header
                .attr("class", "header")
                .append("img")
                .attr(
                    "src",
                    validDataUrl(timeline.HeaderImage) ? timeline.HeaderImage : ""
                ).exit().remove();
        } else if (this.settings.dataPoint.layout.toLowerCase() === "footer") {
            this.footer
                .attr("class", "footer")
                .append("img")
                .attr(
                    "src",
                    validDataUrl(timeline.FooterImage) ? timeline.FooterImage : ""
                );
        }
    }

    private renderXandYAxis(minDate, maxDate, gWidth, gHeight) {
        let xAxis;
        this.xScale = d3.scaleTime()
            .domain([minDate, maxDate])
            .range([this.margin.left, gWidth]);

        if (this.diff_years(minDate, maxDate) <= 1) {
            xAxis = d3.axisBottom(this.xScale)
                .ticks(d3.timeMonth, 1)
                .tickPadding(20)
                .tickFormat(d3.timeFormat("%b '%y"))
                .tickSize(-10);
        }
        else {
            xAxis = d3.axisBottom(this.xScale)
                .ticks(d3.timeYear, 1)
                .tickPadding(20)
                .tickFormat(d3.timeFormat('%Y'))
                .tickSize(-10);
        }

        let xAxisAllTicks = d3.axisBottom(this.xScale)
            .ticks(d3.timeMonth, 3)
            .tickPadding(20)
            .tickFormat(d3.timeFormat(""))
            .tickSize(10);

        this.yScale = d3.scaleLinear()
            .domain([-100, 100])
            .range([gHeight, this.margin.top]);

        let yAxis = d3.axisLeft(this.yScale);

        let xAxisLine = this.svg.append("g")
            .attr("class", "x-axis")
            .attr("transform", "translate(" + (20) + "," + ((gHeight / 2) + 25) + ")")
            .call(xAxis);

        let xAxisLineAllTicks = this.svg.append("g")
            .attr("class", "x-axis")
            .attr("transform", "translate(" + (20) + "," + ((gHeight / 2) + 25) + ")")
            .call(xAxisAllTicks);

        this.svg.append("g")
            .attr("class", "y-axis")
            .call(yAxis).attr('display', 'none');
    }

    private renderTitle(vpWidth, gWidth) {
        let gTitle = this.svg.append('g')
            .attr('x', 0)
            .attr('y', 0)
            .attr('width', vpWidth)
            .attr('height', 35);

        gTitle.append('rect')
            .attr('class', 'chart-header')
            .attr('width', vpWidth)
            .attr('height', 35);

        gTitle.append('text')
            .text('Key Events Timeline')
            .attr('fill', '#ffffff')
            .attr('font-size', 24)
            .attr('transform', 'translate(' + ((gWidth + 70) / 2 - 104) + ',25)');
    }

    private renderLine(timelineData: TimelineData[], gHeight) {
        let _self = this;
        this.svg.selectAll(".line")
            .data(timelineData)
            .enter()
            .append("rect")
            .attr("title", (d) => {
                return sanitizeHtml(d.Description) + '(' + d.Date + ')';
            })
            .attr("x", (d, i) => {
                return _self.xScale(d.Date) + 20;
            })
            .attr("width", '4px')
            .attr("y", (d, i) => {
                if (i % 2 === 0) {
                    return _self.yScale(-27);
                } else {
                    let count = Math.ceil(i / 2);
                    if (count % 2 === 0) {
                        return _self.yScale(70);
                    } else {
                        return _self.yScale(30);
                    }
                }
            })
            .attr("height", (d, i) => {
                if (i % 2 === 0) {
                    let count = i / 2;
                    if (count % 2 === 0) {
                        return gHeight - _self.yScale(-45);
                    }
                    else {
                        return gHeight - _self.yScale(-85);
                    }
                } else {
                    let count = Math.ceil(i / 2);
                    if (count % 2 === 0) {
                        return gHeight - _self.yScale(-45);
                    }
                    else {
                        return gHeight - _self.yScale(-85);
                    }
                }
            })
            .style('fill', (d) => {
                if (d.Type === 'Regulatory') {
                    return this.settings.legendColors.Regulatory;
                }
                else if (d.Type === 'Commercial') {
                    return this.settings.legendColors.Commercial;
                }
                else if (d.Type === 'Clinical Trials') {
                    return this.settings.legendColors.ClinicalTrail;
                }
            });
    }

    private renderArrow(timelineData: TimelineData[]) {
        let _self = this;
        let triangle = d3.symbol().type(d3.symbolTriangle).size(150);

        this.svg.selectAll(".arrow")
            .data(timelineData)
            .enter()
            .append("path")
            .attr('d', triangle)
            .attr("title", (d) => {
                return sanitizeHtml(d.Description) + '(' + d.Date + ')';
            })
            .attr("width", 100)
            .attr("height", 70)
            .style('fill', (d) => {
                if (d.Type === 'Regulatory') {
                    return this.settings.legendColors.Regulatory;
                }
                else if (d.Type === 'Commercial') {
                    return this.settings.legendColors.Commercial;
                }
                else if (d.Type === 'Clinical Trials') {
                    return this.settings.legendColors.ClinicalTrail;
                }
            })
            .attr('transform', (d, i) => {
                let yscale, rotate = '';
                if ((i % 2) === 0) {
                    yscale = _self.yScale(-25);
                } else {
                    yscale = _self.yScale(14);
                    rotate = 'rotate(180)';
                }
                return 'translate(' + (_self.xScale(d.Date) + 22) + ' ' + yscale + ') ' + rotate;
            });
    }

    private renderBox(timelineData: TimelineData[], gWidth, gHeight) {
        let _self = this;
        let gbox = this.svg.selectAll(".box")
            .data(timelineData)
            .enter()
            .append("g")
            .attr('class', (d, i) => {
                if (d.Type === 'Regulatory') { return 'rect regulatory'; } 
                else if (d.Type === 'Commercial') { return 'rect commercial'; }
                else if (d.Type === 'Clinical Trials') { return 'rect clinical-trails'; }
            })
            .style('fill', (d) => {
                if (d.Type === 'Regulatory') { return this.settings.legendColors.Regulatory; }
                else if (d.Type === 'Commercial') { return this.settings.legendColors.Commercial; }
                else if (d.Type === 'Clinical Trials') { return this.settings.legendColors.ClinicalTrail; }
            })
            .attr("title", (d) => { return sanitizeHtml(d.Description) + '(' + d.Date + ')'; })
            .attr("width", () => { return 100; })
            .attr("height", () => { return 70; })
            .attr('transform', (d, i) => { 
                let y;
                if ((i % 2) === 0) {
                    let count = i / 2;
                    if (count % 2 === 0) {
                        y = _self.yScale(-80);
                    } else {
                        y = _self.yScale(-40);
                    }
                } else {
                    let count = Math.ceil(i / 2);
                    if (count % 2 === 0) {
                        y = _self.yScale(95);
                    } else {
                        y = _self.yScale(55);
                    }
                }
                return 'translate(' + (_self.xScale(d.Date) - 25) + ' ' + y + ')';
            });

        gbox.append('a')
            .attr('xlink:href', (d, i) => { return d.DocumentLink; })
            .attr("target", "_blank")
            .append("rect")
            .attr("width", () => { return 100; })
            .attr("height", () => { return 70; })
            .on('click', (e) => { _self.host.launchUrl(e.DocumentLink); });

        gbox.append('a')
            .attr('xlink:href', (d, i) => { return d.DocumentLink; })
            .attr("target", "_blank")
            .append("text")
            .text((d) => { return _self.extractContent(sanitizeHtml(d.Description)); })
            .attr('x', '5')
            .attr('y', '0')
            .attr('fill', '#ffffff')
            .attr('transform', 'translate(10,30)')
            .call(this.wrap, 90)
            .on('click', (e) => { _self.host.launchUrl(e.DocumentLink); });
        
            this.tooltip
            .style("opacity", 0)
            .attr("class", "tooltip")
            .style("position", "absolute")
            .style("background-color", "white")
            .style("border", "solid")
            .style("border-width", "2px")
            .style("border-radius", "5px")
            .style("padding", "5px");

            let self = this;

        gbox.on('mouseover', (d) => {
            self.tooltip.style("opacity", 1);
        })
        .on('mousemove', (d: TimelineData) => {
            var html = '';
                var dateHtml = '<div>' + (d.Date.getMonth() + 1) + '/' + (d.Date.getDate()) + '/' + (d.Date.getFullYear()) + '</div>';
                if (d.Type === "Regulatory") {
                  html = dateHtml + '<div>' + sanitizeHtml(d.Company) + '</div>' + '<div>' + sanitizeHtml(d.Description) + '</div>';
                }
                else if (d.Type === "Commercial") {
                  html = dateHtml + '<div>' + sanitizeHtml(d.Company) + '</div>' + '<div>' + sanitizeHtml(d.Description) + '</div>';
                } else if (d.Type === "Clinical Trials") {
                  html = dateHtml + '<div>' + sanitizeHtml(d.Company) + '</div>' + '<div>' + + sanitizeHtml(d.Description) + '</div>';
                }
                self.tooltip.html(html).style("left", (d3.event.pageX + 20) + "px").style("top", (d3.event.pageY) + "px");
        })
        .on('mouseleave', (d) => {
            self.tooltip.style("opacity", 0);
        });

        gbox.on("mouseenter", function () { d3.select(this).raise(); });

        this.renderLegend(gWidth, gHeight);
    }

    private renderLegend(gWidth, gHeight) {
        let gLegend = this.svg.append('g')
            .attr('transform', 'translate(' + ((gWidth / 2) - 200) + ',' + (gHeight + 45) + ')')
            .attr('x', 0)
            .attr('y', 0)
            .attr('width', gWidth)
            .attr('height', 50);

        let legendClinical = gLegend.append('g')
            .attr('transform', 'translate(50,0)');

        let legendRegulatory = gLegend.append('g')
            .attr('transform', 'translate(200,0)');

        let legendCommercial = gLegend.append('g')
            .attr('transform', 'translate(350,0)');

        gLegend.append('text')
            .text('Code:')
            .attr('transform', 'translate(0,35)');


        legendClinical.append("rect")
            .attr("width", () => {
                return 35;
            })
            .attr("height", () => {
                return 35;
            })
            .attr('fill', this.settings.legendColors.ClinicalTrail);

        legendClinical.append('text')
            .text('Clinical Trials')
            .attr('transform', 'translate(45,35)');


        legendRegulatory.append("rect")
            .attr("width", () => {
                return 35;
            })
            .attr("height", () => {
                return 35;
            })
            .attr('fill', this.settings.legendColors.Regulatory);

        legendRegulatory.append('text')
            .text('Regulatory')
            .attr('transform', 'translate(45,35)');

        legendCommercial.append("rect")
            .attr("width", () => {
                return 35;
            })
            .attr("height", () => {
                return 35;
            })
            .attr('fill', this.settings.legendColors.Commercial);

        legendCommercial.append('text')
            .text('Commercial')
            .attr('transform', 'translate(45,35)');
    }

    public extractContent(str: any) {
        if (str === null || str === "") return false;
        else str = str.toString();
        return str.replace(/(<([^>]+)>)/gi, "");
    }

    public static CONVERTER(dataView: DataView, host: IVisualHost): TimelineData[] {
        let resultData: TimelineData[] = [];
        let tableView = dataView.table;
        let _rows = tableView.rows;
        let _columns = tableView.columns;
        let _companyIndex = -1, _typeIndex = -1, _descIndex = -1, _dateIndex = -1, _linkIndex = -1, _headerImageIndex = -1, _footerImageIndex = -1;
        for (let ti = 0; ti < _columns.length; ti++) {
            if (_columns[ti].roles.hasOwnProperty("Company")) {
                _companyIndex = ti;
            } else if (_columns[ti].roles.hasOwnProperty("Type")) {
                _typeIndex = ti;
            } else if (_columns[ti].roles.hasOwnProperty("Description")) {
                _descIndex = ti;
            } else if (_columns[ti].roles.hasOwnProperty("Date")) {
                _dateIndex = ti;
            } else if (_columns[ti].roles.hasOwnProperty("DocumentLink")) {
                _linkIndex = ti;
            } else if (_columns[ti].roles.hasOwnProperty("HeaderImage")) {
                _headerImageIndex = ti;
            } else if (_columns[ti].roles.hasOwnProperty("FooterImage")) {
                _footerImageIndex = ti;
            }
        }
        for (let i = 0; i < _rows.length; i++) {
            let row = _rows[i];
            let dp = {
                Company: row[_companyIndex].toString(),
                Type: row[_typeIndex] ? row[_typeIndex].toString() : '',
                Description: row[_descIndex] ? row[_descIndex].toString() : null,
                Date: row[_dateIndex] ? new Date(Date.parse(row[_dateIndex].toString())) : null,
                DocumentLink: row[_linkIndex] ? row[_linkIndex].toString() : null,
                HeaderImage: row[_headerImageIndex] ? row[_headerImageIndex].toString() : null,
                FooterImage: row[_footerImageIndex] ? row[_footerImageIndex].toString() : null,
                selectionId: host.createSelectionIdBuilder().withTable(tableView, i).createSelectionId(),
            };
            resultData.push(dp);
        }
        return resultData;
    }

    private static parseSettings(dataView: DataView): VisualSettings {
        return VisualSettings.parse(dataView);
    }

    /**
     * This function gets called for each of the objects defined in the capabilities files and allows you to select which of the
     * objects and properties you want to expose to the users in the property pane.
     *
     */
    public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstance[] | VisualObjectInstanceEnumerationObject {
        return VisualSettings.enumerateObjectInstances(this.settings || VisualSettings.getDefault(), options);
    }

    private wrap(text, width) {
        text.each(function () {
            let text = d3.select(this),
                words = text.text().split(/\s+/).reverse(),
                word,
                line = [],
                lineNumber = 0,
                lineHeight = 1.1,
                x = text.attr("x"),
                y = text.attr("y"),
                dy = 0,
                tspan = text.text(null)
                    .append("tspan")
                    .attr("x", x)
                    .attr("y", y)
                    .attr("dy", dy + "em");
            while (word = words.pop()) {
                line.push(word);
                tspan.text(line.join(" "));
                if (tspan.node().getComputedTextLength() > width) {
                    line.pop();
                    tspan.text(line.join(" "));
                    line = [word];
                    tspan = text.append("tspan")
                        .attr("x", x)
                        .attr("y", y)
                        .attr("dy", ++lineNumber * lineHeight + dy + "em")
                        .text(word);
                }
            }
        });
    }

    private diff_years(dt2, dt1) {
        let diff = (dt2.getTime() - dt1.getTime()) / 1000;
        diff /= (60 * 60 * 24);
        return Math.abs(Math.round(diff / 365.25));
    }
}