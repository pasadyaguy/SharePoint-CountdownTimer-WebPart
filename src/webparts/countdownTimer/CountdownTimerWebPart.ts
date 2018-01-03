import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneButton
} from "@microsoft/sp-webpart-base";
import { escape } from "@microsoft/sp-lodash-subset";
import * as $ from "jquery";
require("./countdown.js");

import styles from "./CountdownTimerWebPart.module.scss";
import * as strings from "CountdownTimerWebPartStrings";

export interface ICountdownTimerWebPartProps {
  description: string;
  eventdate: string;
}

export default class CountdownTimerWebPart extends BaseClientSideWebPart<ICountdownTimerWebPartProps> {
  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.countdownTimer }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
            <div id="clockdiv" class="${ styles.clockdiv }">
              <div>
                <span class="days"></span>
                <div class="${ styles.smalltext }">Days</div>
              </div>
              <div>
                <span class="hours"></span>
                <div class="${ styles.smalltext }">Hours</div>
              </div>
              <div>
                <span class="minutes"></span>
                <div class="${ styles.smalltext }">Minutes</div>
              </div>
              <div>
                <span class="seconds"></span>
                <div class="${ styles.smalltext }">Seconds</div>
              </div>
            </div>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
            </div>
          </div>
        </div>
      </div>`;

      this.startCountdown();
  }

  protected startCountdown(): any {
    this.initializeClock("clockdiv", this.properties.eventdate);
  }

  protected initializeClock(id: string, endtime: string): any {
    let clock: HTMLElement = document.getElementById(id);
    let daysSpan: Element = clock.querySelector(".days");
    let hoursSpan: Element = clock.querySelector(".hours");
    let minutesSpan: Element = clock.querySelector(".minutes");
    let secondsSpan: Element = clock.querySelector(".seconds");

    function getTimeRemaining(endtime: any): any {
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
      let t: any = getTimeRemaining(endtime);

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
                PropertyPaneTextField("eventdate", {
                  label: "Event Date"
                }),
                PropertyPaneTextField("description", {
                  label: "Event Description"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
