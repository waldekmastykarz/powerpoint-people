/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import { IDynamicPerson, Providers, ProviderState, SimpleProvider } from "@microsoft/mgt";
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global $, document, Office */

import { getAccessToken, signIn } from "./../helpers/ssoauthhelper";

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    $(document).ready(function () {
      $("#signin").click(signIn);
    });
  }
});

export function writeDataToOfficeDocument(result: Object): void {
  let data: string[] = [];
  let userProfileInfo: string[] = [];
  userProfileInfo.push(result["displayName"]);
  userProfileInfo.push(result["jobTitle"]);
  userProfileInfo.push(result["mail"]);
  userProfileInfo.push(result["mobilePhone"]);
  userProfileInfo.push(result["officeLocation"]);

  for (let i = 0; i < userProfileInfo.length; i++) {
    if (userProfileInfo[i] !== null) {
      data.push(userProfileInfo[i]);
    }
  }

  let userInfo: string = "";
  for (let i = 0; i < data.length; i++) {
    userInfo += data[i] + "\n";
  }
  Office.context.document.setSelectedDataAsync(userInfo, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      throw asyncResult.error.message;
    }
  });
}

let imageCount = 0;

export function showPeoplePicker(): void {
  Providers.globalProvider = new SimpleProvider(
    () => Promise.resolve(getAccessToken()),
    () => {
      Providers.globalProvider.setState(ProviderState.SignedIn);
      return Promise.resolve();
    },
    function () {
      Providers.globalProvider.setState(ProviderState.SignedOut);
      return Promise.resolve();
    }
  );
  Providers.globalProvider.setState(ProviderState.SignedIn);
  $("#signin").remove();
  $("#content").append('<mgt-people-picker type="Person"></mgt-people-picker><div id="people"></div>');
  $("mgt-people-picker")[0].addEventListener("selectionChanged", () => {
    $("#people").html(`<mgt-people show-max="100">
    <template>
      <ul style="padding-left: 0">
        <li data-for="person in people" style="list-style-type: none">
          <mgt-person person-details="{{person}}" fetch-image="true" view="twolines" line2-property="jobTitle" data-props="{{@click: personClick}}">
          </mgt-person>
        </li>
      </ul>
    </template>
  </mgt-people>
  <button>Insert all</button>`);
    ($("mgt-people")[0] as any).templateContext = {
      personClick: (e, context) => {
        if (clickedOnPhoto(e)) {
          addPicture(context.person, true);
        } else {
          addInfo(context.person);
        }
      },
    };
    ($("mgt-people")[0] as any).people = [
      ...(($("#content mgt-people-picker")[0] as any).selectedPeople as IDynamicPerson[]),
    ];
    $("#content button").click(() => {
      imageCount = 0;
      const people = [...(($("#content mgt-people-picker")[0] as any).selectedPeople as IDynamicPerson[])];
      addPeopleInfo(people);
    });
  });
}

function clickedOnPhoto(event: Event) {
  const path = (event as any).path;
  if (!path) {
    return false;
  }

  return path[2].className === "user-avatar" || path[0].className.indexOf("initials") > -1;
}

function addPicture(person: IDynamicPerson, single: boolean) {
  return new Promise<void>((resolve, reject): void => {
    if (!person.personImage) {
      return resolve();
    }
    const options: Office.SetSelectedDataOptions = {
      coercionType: Office.CoercionType.Image,
      imageWidth: 100,
    };
    if (!single) {
      options.imageLeft = 150 + 150 * imageCount;
      options.imageTop = 400;
    }
    Office.context.document.setSelectedDataAsync(
      person.personImage.substr(person.personImage.indexOf(",") + 1),
      options,
      (res) => {
        if (res.error) {
          return reject(res.error.message);
        }

        resolve();
      }
    );
  });
}

function addInfo(person: IDynamicPerson) {
  return new Promise<void>((resolve, reject): void => {
    const p = `${person.displayName}\n${person.jobTitle}`;
    Office.context.document.setSelectedDataAsync(p, (res) => {
      if (res.error) {
        return reject(res.error.message);
      }

      resolve();
    });
  });
}

function addPeopleInfo(people: IDynamicPerson[]) {
  const person = people.shift();
  addPersonInfo(person).then(
    () => {
      if (people.length > 0) {
        imageCount++;
        addPeopleInfo(people);
      }
    },
    (err) => logError(err)
  );
}

function addPersonInfo(person: IDynamicPerson): Promise<void> {
  return new Promise<void>((resolve, reject): void => {
    addInfo(person).then(
      () => {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            return reject("Action failed with error: " + asyncResult.error.message);
          }

          const currentSlideId = (asyncResult.value as any).slides[0].id;
          Office.context.document.goToByIdAsync(currentSlideId, Office.GoToType.Slide, function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              return reject("Action failed with error: " + asyncResult.error.message);
            }

            if (!person.personImage) {
              return resolve();
            }

            addPicture(person, false).then(() => {
              Office.context.document.goToByIdAsync(currentSlideId, Office.GoToType.Slide, function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                  return reject("Action failed with error: " + asyncResult.error.message);
                }

                resolve();
              });
            });
          });
        });
      },
      (err) => logError(err)
    );
  });
}

function logError(error: string) {
  console.error(error);
}
