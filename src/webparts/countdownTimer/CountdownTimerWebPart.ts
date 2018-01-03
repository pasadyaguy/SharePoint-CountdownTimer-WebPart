import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneButton,
  PropertyPaneSlider
} from "@microsoft/sp-webpart-base";
import { escape } from "@microsoft/sp-lodash-subset";
import * as $ from "jquery";

import styles from "./CountdownTimerWebPart.module.scss";
import * as strings from "CountdownTimerWebPartStrings";

export interface ICountdownTimerWebPartProps {
  description: string;
  eventdate: string;
  eventname: string;
  paddingsize: string;
}

export default class CountdownTimerWebPart extends BaseClientSideWebPart<ICountdownTimerWebPartProps> {
  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.countdownTimer }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <div id="${ this.properties.eventname }" class="${ styles.clockdiv }">
                <div style="padding: ${ this.properties.paddingsize }px">
                  <span class="days"></span>
                  <div class="${ styles.smalltext }">Days</div>
                </div>
                <div style="padding: ${ this.properties.paddingsize }px">
                  <span class="hours"></span>
                  <div class="${ styles.smalltext }">Hours</div>
                </div>
                <div style="padding: ${ this.properties.paddingsize }px">
                  <span class="minutes"></span>
                  <div class="${ styles.smalltext }">Minutes</div>
                </div>
                <div style="padding: ${ this.properties.paddingsize }px">
                  <span class="seconds"></span>
                  <div class="${ styles.smalltext }">Seconds</div>
                </div>
              </div>
              <div class="${ styles.description }">${ escape(this.properties.description) }</div>
            </div>
          </div>
        </div>
      </div>`;

      this.startCountdown();
  }

  protected startCountdown(): any {
    this.initializeClock(this.properties.eventname, this.properties.eventdate);
  }

  protected initializeClock(id: string, endtime: string): any {
    let clock: HTMLElement = document.getElementById(id);
    let daysSpan: Element = clock.querySelector(".days");
    let hoursSpan: Element = clock.querySelector(".hours");
    let minutesSpan: Element = clock.querySelector(".minutes");
    let secondsSpan: Element = clock.querySelector(".seconds");

    function getTimeRemaining(): any {
      let now: any = new Date();
      let t: any = Date.parse(endtime) - now;
      let seconds: number = Math.floor((t / 1000) % 60);
      let minutes: number = Math.floor((t / 1000 / 60) % 60);
      let hours: number = Math.floor((t / (1000 * 60 * 60)) % 24);
      let days: number = Math.floor(t / (1000 * 60 * 60 * 24));
      return {
        "total": t,
        "days": days,
        "hours": hours,
        "minutes": minutes,
        "seconds": seconds
      };
    }

    function updateClock(): any {
      let t: any = getTimeRemaining();

      daysSpan.innerHTML = t.days;
      hoursSpan.innerHTML = ("0" + t.hours).slice(-2);
      minutesSpan.innerHTML = ("0" + t.minutes).slice(-2);
      secondsSpan.innerHTML = ("0" + t.seconds).slice(-2);
    }

    updateClock();
    setInterval(() => updateClock(), 1000);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("eventname", {
                  label: "Event Name"
                }),
                PropertyPaneTextField("eventdate", {
                  label: "Event Date"
                }),
                PropertyPaneTextField("description", {
                  label: "Event Description"
                })
              ]
            },
            {
              groupName: "Appearence",
              groupFields: [
                PropertyPaneSlider("paddingsize", {
                  label: "Size",
                  min: 0,
                  max: 10,
                  showValue:true
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
