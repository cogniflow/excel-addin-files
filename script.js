const STATUS_PENDING = 0; //pending
const STATUS_IN_QUEUE = 1; //waiting
const STATUS_RUNNING = 2; //running
const STATUS_FINISHED = 3; //ready
const STATUS_CANCELED_BY_USER = 4; //canceled
const STATUS_CANCELED_BY_TIMEOUT = 5; //timed out
const STATUS_FINISHED_WITH_ERROR = 6; //error

const QANDA_TASK = 3;

const EXPERIMENT = {
  current: {},
  models: [
    { name: 'FFN Vectorization algorithm' },
    { name: 'Linear Regression' },
  ],
  currentModel: {},
  query: 'sort=recent',
  list: [],
  showConfidence: false,
  insertPlace: 'insertRight',
  colRange: 'A2:A5',
  hasHeaders: true,
  publicExperiments: [],
  visibilityQuery: '',
};
const API_URL = 'https://api.cogniflow.ai';
const PREDICT_API_URL = 'https://predict.cogniflow.ai';
const logoBaseUrl = 'https://s3.us-east-2.amazonaws.com/static-cogniflow-prod/';

const EXPERIMENT_TASKS = {
  0: 'Text classification',
  1: 'Text Translation',
  2: 'Entities recognition',
  3: 'Question and answer',
  100: 'Image classification',
  200: 'Audio classification',
  201: 'Audio speech to text',
};

const EXPERIMENT_TYPES = {
  0: 'Text',
  1: 'Image',
  2: 'Audio',
};

const EXPERIMENT_TYPES_URL = {
  0: 'text',
  1: 'image',
  2: 'audio',
};

const URL_TASK_MAP = {
  0: 'classification/predict',
  1: 'translation/translate',
  2: 'information-extraction/extract-entities',
  3: 'question-answering/ask',
  100: 'classification/predict-from-web',
  200: 'classification/predict-from-web',
  201: 'classification/predict-from-web',
};
// Task based
const REQUEST_BODY_KEYS = {
  0: 'text',
  3: 'question',
};

let user = {};

const pages = {
  login: {
    pageId: 'page-login',
    goToSecretTokenAuth: 'login-go-to-secret',
    gotToNormalAuth: 'login-go-to-normal-auth',
    subPages: {
      defaultSubpage: 'login-default-form',
      secretToken: 'login-use-secret-token',
    },
    emailInput: 'login-email-field',
    passwordInput: 'login-password-field',
    loginButton: 'default-login-button',
    invalidCredentials: 'login-invalid-creds',
  },
  experiments: {
    pageId: 'page-experiments',
    experimentCardsClassName: 'experiment-card',
    cardsContainer: 'experiment-cards-container',
    visibilityDropdown: 'experiments-visibility-dropdown',
    typeDropdown: 'experiments-type-dropdown',
    logoutBtn: 'logout',
  },
  run: {
    pageId: 'pages-run-model',
    goBack: 'pages-run-model-go-back',

    confidenceDd: 'run-show-confidence-checkbox',
    includeHeadersContainer: 'run-include-headers-container',
    hasHeadersCb: 'run-show-has-headers',
    insertDd: 'run-insert-dropdown',
    modelDd: 'run-model-dropdown',

    cellsRange: 'run-cells-range',

    expValue: 'experiment-value',
    confidenceValue: 'confidence-value',
    hasHeadersValue: 'has-headers-value',
    modelValue: 'model-value',
    resultPlaceValue: 'result-place-value',

    runProgressContainer: 'run-progress-container',
    runProcessed: 'run-processed',
    runTotalCells: 'run-total-cells',
    runPercentage: 'run-percentage',
    runStop: 'run-stop',

    runExperiment: 'run-experiment',
    runEditWarning: 'run-edit-warning',
  },
};

const runState = {
  progress: 0,
  processed: 0,
  total: 0,

  stop: false,

  updateProgress() {
    this.progress = parseInt((this.processed / this.total) * 100);
  },
  clearRunState(dom) {
    this.progress = 0;
    this.processed = 0;
    this.total = 0;

    this.stop = false;

    dom.changeText(pages.run.runStop, 'Stop');

    dom.changeText(pages.run.runTotalCells, this.total);
    dom.changeText(pages.run.runProcessed, this.processed);
    dom.changeText(pages.run.runPercentage, this.progress);
  },
};

// const Excel = {
//   run: () => { },
// };

const EXCEL_EVENTS = {
  select: null,
};

const formattedDate = (dateStr) => {
  const date = new Date(dateStr);
  const day = date.getDate();
  const month = date.getMonth() + 1;
  const year = date.getFullYear();

  if (month < 10) {
    return `${day}/0${month}/${year}`;
  } else {
    return `${day}/${month}/${year}`;
  }
};
// https://app.dev.cogniflow.ai/static/media/img-placeholder.bf7d98c5.svg
const logoPlaceholder =
  'https://app.dev.cogniflow.ai/static/media/img-placeholder.bf7d98c5.svg';

const iconSvgExperimentTypeMap = {
  text: `<svg width="32" height="32" viewBox="0 0 32 32" style="display:flex;align-items:center;background-color:#eceffa;border-radius:8px;"><g fill="none" fill-rule="evenodd"><path fill="#a4a9c8" d="M24 21h-6c-.265 0-.52.107-.707.294l-.293.293V11.415L18.414 10H24v11zm-10 0H8V10h5.586L15 11.416v10.172l-.293-.293C14.52 21.106 14.265 21 14 21zM25 8h-7c-.265 0-.52.107-.707.293L16 9.586l-1.293-1.293c-.187-.186-.442-.292-.707-.292H7c-.552 0-1 .447-1 1v13c0 .551.448 1 1 1h6.586l1.707 1.706c.195.195.451.293.707.293.256 0 .512-.098.707-.293L18.414 23H25c.552 0 1-.448 1-1V9c0-.552-.448-1-1-1z"></path></g></svg>`,
  image: `<svg width="32" height="33" viewBox="0 0 32 33" style="display:flex;align-items:center;background-color:#eceffa;border-radius:8px;"><g fill="#A4A9C8" fill-rule="evenodd"><path d="M2 18v-1.52l3.93-3.14 2.36 2.37c.366.367.953.394 1.35.06l5.3-4.42L18 14.41V18H2zM18 2v9.59l-2.29-2.3c-.366-.367-.952-.393-1.35-.06l-5.3 4.42-2.35-2.36c-.361-.358-.933-.388-1.33-.07L2 13.92V2h16zm1-2H1C.448 0 0 .448 0 1v18c0 .553.448 1 1 1h18c.553 0 1-.447 1-1V1c0-.552-.447-1-1-1z" transform="translate(6 7)"></path><path d="M9 6c.552 0 1 .448 1 1s-.448 1-1 1-1-.448-1-1 .448-1 1-1m0 4c1.657 0 3-1.343 3-3s-1.343-3-3-3-3 1.343-3 3 1.343 3 3 3" transform="translate(6 7)"></path></g></svg>`,
  audio: `<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" style="display:flex;align-items:center;border-radius:8px;background-color:#eceffa;fill: #A4A9C8;transform: ;msFilter:;"><path d="M8 12c2.28 0 4-1.72 4-4s-1.72-4-4-4-4 1.72-4 4 1.72 4 4 4zm0-6c1.178 0 2 .822 2 2s-.822 2-2 2-2-.822-2-2 .822-2 2-2zm1 7H7c-2.757 0-5 2.243-5 5v1h2v-1c0-1.654 1.346-3 3-3h2c1.654 0 3 1.346 3 3v1h2v-1c0-2.757-2.243-5-5-5zm9.364-10.364L16.95 4.05C18.271 5.373 19 7.131 19 9s-.729 3.627-2.05 4.95l1.414 1.414C20.064 13.663 21 11.403 21 9s-.936-4.663-2.636-6.364z"></path><path d="M15.535 5.464 14.121 6.88C14.688 7.445 15 8.198 15 9s-.312 1.555-.879 2.12l1.414 1.416C16.479 11.592 17 10.337 17 9s-.521-2.592-1.465-3.536z"></path></svg>`,
};

const experimentTemplate = ({
  id,
  title,
  type,
  task,
  created_at,
  logo,
  ...rest
}) => {
  const typeString = EXPERIMENT_TYPES[type];
  const taskString = EXPERIMENT_TASKS[task];
  const logoUrl = logo ? `${logoBaseUrl}${logo}` : logoPlaceholder;

  return `
  <div class="experiment-card" id="${id}">

    <div class="exp-card-header">
      <div class="exp-card-bg-img"></div>
      <div class="exp-card-header-logos">
        <div class="exp-card-logo-wrapper">
          <img src="${logoUrl}" class="exp-card-logo" alt="card logo">
        </div>
        <div class="exp-card-logo-type-wrapper">
          ${iconSvgExperimentTypeMap[EXPERIMENT_TYPES_URL[type]]}
        </div>

      </div>
    </div>
    
    <div class="exp-card-content">

      <div class="exp-card-title">
        <p title="${title}">${title}</p>
      </div>
      <div class="exp-card-body">
        <p>${typeString}/${taskString}</p>
      </div>
      <div class="exp-card-time">
        <p>
          Created: <span class="exp-card-time-label">${formattedDate(
            created_at
          )}</span>
        </p>
      </div>

    </div>
  </div>
  `;
};

const optionTemplate = ({ id, name, recommended }) => {
  const icon = recommended ? '✓' : '';

  return `
    <option value="${name}" id={${id}}>${name} ${icon}</option>
  `;
};

const setExperimentsInDom = (
  experiments,
  templateFunc,
  experimentsParentId,
  domHandler
) => {
  const container = domHandler.getById(experimentsParentId);
  let allExperimentsTemplate = '';

  for (const exp of experiments) {
    allExperimentsTemplate += templateFunc(exp);
  }
  container.innerHTML = allExperimentsTemplate;
};

const setModelOptionsInDom = (models, idSelect, domHandler) => {
  const domSelectElement = domHandler.getById(idSelect);
  let allModels = '';

  for (const model of models) {
    allModels += optionTemplate(model);
  }
  domSelectElement.innerHTML = allModels;
};

const setRecommendedModel = () => {
  if (!EXPERIMENT.current.id_recommended_model) {
    EXPERIMENT.currentModel = {
      recommended: true,
      ...EXPERIMENT.models[0],
    };
    return;
  }

  const recommendedIdx = EXPERIMENT.models.findIndex(
    (model) => EXPERIMENT.current.id_recommended_model === model.id
  );

  EXPERIMENT.currentModel = {
    recommended: true,
    ...EXPERIMENT.models[recommendedIdx],
  };
  EXPERIMENT.models[recommendedIdx] = EXPERIMENT.models[0];
  EXPERIMENT.models[0] = EXPERIMENT.currentModel;
};

const findAndSetModelByName = (modelName) => {
  EXPERIMENT.currentModel = EXPERIMENT.models.find(
    ({ name }) => name === modelName
  );
};

const setColumnAtTheRight = async (
  range,
  sheet,
  context,
  isConfidence = false
) => {
  const nextEntireCol = range
    .getEntireColumn()
    .getColumnsAfter(1)
    .insert('Right');

  nextEntireCol.load('address');
  nextEntireCol.load('values');

  await context.sync();

  const newColChar = nextEntireCol.address.split('!')[1][0];
  const [startCell, endCell] = range.address.split('!')[1].split(':');
  const [_, startIdx] = startCell;
  const endIdx = endCell ? endCell[1] : '';

  const insertedColumnRangeStr = endIdx
    ? `${newColChar}${startIdx}:${newColChar}${endIdx}`
    : `${newColChar}${startIdx}`;
  const newInsertedColumnRange = sheet.getRange(insertedColumnRangeStr);

  if (EXPERIMENT.hasHeaders) {
    if (isConfidence) {
      const titleCellRange = sheet.getRange(`${newColChar}1`);
      titleCellRange.values = [['Confidence']];
    } else {
      const titleCellRange = sheet.getRange(`${newColChar}1`);
      titleCellRange.values = [[EXPERIMENT.current.title]];
    }
  }

  newInsertedColumnRange.load('values');
  newInsertedColumnRange.load('address');

  console.log(insertedColumnRangeStr);

  await context.sync();

  return newInsertedColumnRange;
};

const columnToPlaceResultMap = {
  insertRight: (range) => range.getColumnsAfter(1),
  replaceRight: (range) => range.getColumnsAfter(1),
};

const isEntireColumn = (address) => {
  const [addressStart, addressEnd] = address.split(':');

  return addressStart === addressEnd || addressStart.length === 1;
};

const delayRequest = (time) =>
  new Promise((resolve) => {
    setTimeout(() => {
      resolve();
    }, time);
  });

const getNextFromChar = (char, n = 0) =>
  String.fromCharCode(char.charCodeAt() + n);

const getColumnNameFromRange = (rangeStr) => rangeStr.split('')[0];

const getColmunStartIndex = (rangeStr) => {
  const [startCell] = rangeStr.split(':');
  const [_, ...nums] = startCell.split('');
  const index = nums.join('');

  return parseInt(index);
};

const getEntireColumnRangeString = async (
  sheet,
  context,
  forInsertResult = false
) => {
  const range = sheet.getUsedRange();
  const columnLetter = getColumnNameFromRange(EXPERIMENT.colRange);
  const letterIndex = columnLetter.charCodeAt() - 65;
  const start = EXPERIMENT.hasHeaders ? 1 : 0;
  let endIdx = start;

  range.load('values');
  await context.sync();
  for (endIdx; endIdx < range.values.length; endIdx++) {
    if (!range.values[endIdx][letterIndex]) break;
  }
  return `${columnLetter}${start + 1}:${columnLetter}${endIdx}`;
};

function removeExcelEvent() {
  EXCEL_EVENTS.select.remove();
  return Excel.run(EXCEL_EVENTS.select.context, function (context) {
    return context.sync().then(function () {
      EXCEL_EVENTS.select = null;
      // console.log("Event handler successfully removed.");
    });
  });
}

function logout(router) {}

const testToken = '';
// "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJzdWIiOiJmMDk0Mjg4Yi0wYWFjLTRhYWQtOGRmNi00NjhlZTY4OWEwNGUiLCJhY2NvdW50X3R5cGUiOjMsImV4cCI6MTY0Njg1NjEwM30.Qc222N0PhQjdBAP84aCc979rmoCWkK9AKq1ddnk2yzY";

function wait(delay) {
  return new Promise((resolve) => setTimeout(resolve, delay));
}

function fetchRetry(url, delay, tries, fetchOptions = {}) {
  function onError(err) {
    triesLeft = tries - 1;
    if (!triesLeft) {
      throw err;
    }
    return wait(delay).then(() =>
      fetchRetry(url, delay, triesLeft, fetchOptions)
    );
  }
  return fetch(url, fetchOptions).catch(onError);
}
class HttpService {
  constructor(baseUrl) {
    this.headers = {
      'Content-Type': 'application/json',
    };
    this.baseUrl = baseUrl;

    testToken && this.setTokenInHeaders(testToken);
  }

  setTokenInHeaders(token) {
    this.headers = { Authorization: `Bearer ${token}`, ...this.headers };
  }

  req(method) {
    return (url, body, headers, type = 'normal') => {
      if (type === 'normal') {
        return fetch(`${this.baseUrl}/${url}`, {
          method,
          headers: { ...this.headers, ...(headers ?? {}) },
          body: body,
        })
          .then((res) => {
            if (!res.ok) {
              const error = new Error('HTTP status code: ' + res.status);
              error.response = res;
              error.status = res.status;
              throw error;
            }
            return res.json();
          })
          .then((body) => {
            return body;
          });
      } else {
        return fetchRetry(`${this.baseUrl}/${url}`, 120000, 200, {
          method,
          headers: { ...this.headers, ...(headers ?? {}) },
          body: body,
        })
          .then((res) => {
            if (!res.ok) {
              const error = new Error('HTTP status code: ' + res.status);
              error.response = res;
              error.status = res.status;
              throw error;
            }
            return res.json();
          })
          .then((body) => {
            return body;
          });
      }
    };
  }
  get = this.req('GET');
  post = this.req('POST');
}

class DomHandler {
  getById(id) {
    return document.getElementById(id);
  }

  getElementsByClass(className) {
    return document.getElementsByClassName(className);
  }

  hidePage(pageId) {
    this.getById(pageId).style.display = 'none';
  }
  showPage(pageId) {
    this.getById(pageId).style.display = 'block';
  }

  addEvent(elementId, eventName, callback) {
    this.getById(elementId).addEventListener(eventName, callback);
  }

  addEventToElements(elementsClass, eventName, callback) {
    const elements = this.getElementsByClass(elementsClass);

    for (const element of elements) {
      element.addEventListener(eventName, callback);
    }
  }

  changeText(elementId, text) {
    this.getById(elementId).innerText = text;
  }

  removeListenersFromNode(nodeId) {
    const el = document.getElementById(nodeId);
    const elClone = el.cloneNode(true);

    el.parentNode.replaceChild(elClone, el);
  }
}

class PageStateHandler {
  constructor(initialPageId, domHandler) {
    this.current = initialPageId;
    this.history = [this.current];
    this.dom = domHandler;

    this.dom.showPage(initialPageId);
  }

  updateForwardPageState(newPageId) {
    this.history.push(newPageId);
    this.current = this.history[this.history.length - 1];
  }

  switchPagesVisibility(visible, toShow) {
    this.dom.hidePage(visible);
    this.dom.showPage(toShow);
  }

  navigateTo(newPageId) {
    this.switchPagesVisibility(this.current, newPageId);
    this.updateForwardPageState(newPageId);
  }

  navigateToSubPage(currentSubPage, newSubPage) {
    this.switchPagesVisibility(currentSubPage, newSubPage);
  }
}

async function runExcelAPI(dom) {
  if (EXPERIMENT.colRange.length === 1) {
    const colLetter = getColumnNameFromRange(EXPERIMENT.colRange);

    EXPERIMENT.colRange = `${colLetter}:${colLetter}`;
  }

  const { showConfidence, currentModel, insertPlace } = EXPERIMENT;
  let { colRange } = EXPERIMENT;

  const httpPredict = new HttpService(PREDICT_API_URL);

  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    if (isEntireColumn(colRange)) {
      colRange = await getEntireColumnRangeString(sheet, context);
    } else {
      const selectedRange = sheet.getRange(colRange);

      selectedRange.load('values');
      await context.sync();

      const start = getColmunStartIndex(colRange);
      const columnName = getColumnNameFromRange(colRange);
      let end = start;

      for (const [val] of selectedRange.values) {
        if (!val) break;
        end++;
      }
      if (end !== start) {
        colRange = `${columnName}${start}:${columnName}${end - 1}`;
      }
    }
    const range = sheet.getRange(colRange);

    range.load('values');
    range.load('address');

    await context.sync();

    const requestInterval = {
      text: {
        time: 1000,
        limit: 5,
      },
      image: {
        time: 1000,
        limit: 2,
      },
      audio: {
        time: 1000,
        limit: 1,
      },

      index: 0,
    };
    const column = range.values;
    let requests = [];
    let idx = 0;
    let prev = 0;

    //  -------------------------------------------------------------------------------------

    runState.total = column.length;

    const typeUrl = EXPERIMENT_TYPES_URL[EXPERIMENT.current.type];
    const task = URL_TASK_MAP[EXPERIMENT.current.task];
    const currentModelId = currentModel.id;

    async function fillCellsRequest(text, idx, data) {
      if (runState.stop) return;

      const responses = await Promise.allSettled(data);

      let resultText = [];
      if (EXPERIMENT.current.task === QANDA_TASK) {
        resultText = responses.map(({ value }) => [value.result[0].answer]);
      } else {
        resultText = responses.map(({ value }) => [value.result]);
      }
      const currrentColumnLetter = getColumnNameFromRange(colRange);
      const height = getColmunStartIndex(colRange);
      const responseRangeString = `${currrentColumnLetter}${
        prev + height
      }:${currrentColumnLetter}${idx + height}`;
      let responseRange = sheet.getRange(responseRangeString);

      responseRange.load('values');
      responseRange.load('address');
      await context.sync();

      responseRange = columnToPlaceResultMap[insertPlace](responseRange);

      responseRange.load('values');

      prev = idx + 1;
      responseRange.values = resultText;
      if (showConfidence) {
        let confidenceResponse = 0;
        if (EXPERIMENT.current.task === QANDA_TASK) {
          confidenceResponse = responses.map(({ value }) => [
            value.result[0].confidence,
          ]);
        } else {
          confidenceResponse = responses.map(({ value }) => [
            value.confidence_score,
          ]);
        }
        const confidenceColumn = columnToPlaceResultMap[insertPlace](
          responseRange
        );

        confidenceColumn.values = confidenceResponse;
      }
      requestInterval.index = -1;

      if (!runState.stop) {
        runState.processed = runState.processed + responses.length;
        runState.updateProgress();

        dom.changeText(pages.run.runTotalCells, runState.total);
        dom.changeText(pages.run.runProcessed, runState.processed);
        dom.changeText(pages.run.runPercentage, runState.progress);
      }

      await context.sync();
    }
    const currentType = EXPERIMENT.current.type;
    const fileFormatMap = {
      image: 'jpg',
      audio: 'mp3',
    };
    const format = fileFormatMap[EXPERIMENT_TYPES_URL[currentType]];
    const intervals = requestInterval[EXPERIMENT_TYPES_URL[currentType]];

    if (EXPERIMENT.insertPlace === 'insertRight') {
      if (EXPERIMENT.showConfidence) {
        const newResultCol = await setColumnAtTheRight(
          range,
          sheet,
          context,
          false
        );
        await setColumnAtTheRight(newResultCol, sheet, context, true);
      } else {
        await setColumnAtTheRight(range, sheet, context, false);
      }
    }

    for (const cell of column) {
      if (runState.stop) break;
      const [text] = cell;
      const replaced = text.replace(/\n/g, '\\n');
      const bodyKey = REQUEST_BODY_KEYS[EXPERIMENT.current.task] || 'url';

      requests.push(
        httpPredict.post(
          `${typeUrl}/${task}/${currentModelId}`,
          JSON.stringify({
            [bodyKey]: replaced,
            ...(format ? { format } : {}),
          }),
          {
            accept: 'application/json',
            'Content-Type': 'application/json',
            'x-api-key': user.api_keys ? user.api_keys[0].key : {},
          },
          'retry'
        )
      );

      if (
        requestInterval.index === intervals.limit ||
        idx === column.length - 1
      ) {
        await fillCellsRequest(text, idx, requests);
        requests = [];
      }
      requestInterval.index += 1;
      idx += 1;
    }
    await context.sync();
    dom.getById(pages.run.runEditWarning).style.display = 'none';
    dom.changeText(pages.run.runExperiment, 'Run model');
  });
}

function setUpListeners(dom, router, pages, http) {
  const {
    login: {
      pageId: loginPageId,
      emailInput,
      passwordInput,
      loginButton,
      goToSecretTokenAuth,
      gotToNormalAuth,
      subPages: { defaultSubpage, secretToken },
      invalidCredentials,
    },
    experiments: {
      pageId: expPageId,
      experimentCardsClassName,
      cardsContainer,
      visibilityDropdown,
      typeDropdown,
      logoutBtn,
    },
    run: {
      pageId: runPageId,
      goBack: runPageGoBack,
      confidenceDd,
      includeHeadersContainer,
      hasHeadersCb,
      insertDd,
      modelDd,
      expValue,
      confidenceValue,
      hasHeadersValue,
      modelValue,
      resultPlaceValue,
      cellsRange,

      runProgressContainer,
      runProcessed,
      runTotalCells,
      runPercentage,
      runStop,

      runExperiment,
      runEditWarning,
    },
  } = pages;

  const setLoginEvents = (cb) => {
    dom.addEvent(loginButton, 'click', () => {
      const cred = {
        username: dom.getById(emailInput).value,
        password: dom.getById(passwordInput).value,
      };
      http
        .post('login', new URLSearchParams({ ...cred }), {
          'Content-Type': 'application/x-www-form-urlencoded',
        })
        .then((res) => {
          http.setTokenInHeaders(res.access_token);
          router.navigateTo(expPageId);
          dom.getById(invalidCredentials).style.display = 'none';
          cb();
        })
        .catch((err) => {
          const messages = {
            422: 'Sorry, something went wrong, please try later',
            400: 'Invalid credencials, please try again',
          };
          dom.getById(invalidCredentials).style.display = 'block';
          dom.getById(invalidCredentials).innerText = messages[err.status];
          dom.getById(loginButton).blur();
        });
    });
    dom.addEvent(goToSecretTokenAuth, 'click', () => {
      router.switchPagesVisibility(defaultSubpage, secretToken);
    });
    dom.addEvent(gotToNormalAuth, 'click', () => {
      router.switchPagesVisibility(secretToken, defaultSubpage);
    });
  };

  const setExperimentEvents = (cb) => {
    http
      .get('user/')
      .then((res) => {
        user = res;
      })
      .catch((err) => {
        console.log(err);
      });
    const appendExperimentListInDom = (experimentList) => {
      dom.getById(cardsContainer).innerHTML = '';

      setExperimentsInDom(
        experimentList,
        experimentTemplate,
        cardsContainer,
        dom
      );

      dom.addEventToElements(experimentCardsClassName, 'click', (event) => {
        const selectedExperimentId = event.currentTarget.id;

        http
          .get(`experiment/${selectedExperimentId}`)
          .then(({ models, ...rest }) => {
            EXPERIMENT.current = rest;
            EXPERIMENT.models = models;
            router.navigateTo(runPageId);
            cb();
          });
      });
    };

    http.get(`experiment/?${EXPERIMENT.query}`).then((response) => {
      EXPERIMENT.list = response.filter(
        (experiment) => experiment.status === STATUS_FINISHED
      );
      appendExperimentListInDom(EXPERIMENT.list);
    });

    http.get(`experiment/public?${EXPERIMENT.query}`).then((response) => {
      EXPERIMENT.publicExperiments = response;
      dom.getById(visibilityDropdown).disabled = false;

      dom.addEvent(visibilityDropdown, 'change', (event) => {
        const { value: visibility } = event.target;
        const experimentList = EXPERIMENT[visibility];
        EXPERIMENT.visibilityQuery =
          visibility === 'publicExperiments' ? 'public' : '';

        appendExperimentListInDom(experimentList);
      });
    });

    dom.addEvent(typeDropdown, 'change', (event) => {
      const visibility = EXPERIMENT.visibilityQuery;
      const type = event.target.value;
      const typeQuery = type ? `type=${type}&&` : '';

      http
        .get(`experiment/${visibility}?${typeQuery}${EXPERIMENT.query}`)
        .then((response) => {
          appendExperimentListInDom(response);
        });
    });

    dom.addEvent(logoutBtn, 'click', () => {
      router.navigateTo(loginPageId);
      dom.removeListenersFromNode(expPageId);
    });
  };

  const setRunModelPageEvents = () => {
    setRecommendedModel();
    setModelOptionsInDom(EXPERIMENT.models, modelDd, dom);

    Excel.run((context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      EXCEL_EVENTS.select = sheet.onSelectionChanged.add((event) => {
        dom.getById(cellsRange).value = event.address;
        EXPERIMENT.colRange = event.address;
        // dom.getById(includeHeadersContainer).style.display = isEntireColumn(event.address) ? "flex" : "none";
      });

      return context.sync();
    });

    dom.changeText(expValue, EXPERIMENT.current.title);
    // dom.changeText(confidenceValue, false);
    // dom.changeText(resultPlaceValue, 'insertRight');
    // dom.changeText(modelValue, EXPERIMENT.currentModel.name);

    dom.addEvent(runPageGoBack, 'click', () => {
      dom.changeText(pages.run.runStop, 'Run model');
      runState.clearRunState(dom);
      EXPERIMENT.insertPlace = 'insertRight';
      dom.removeListenersFromNode(runPageId);
      router.navigateTo(expPageId);
      removeExcelEvent();
    });
    dom.addEvent(cellsRange, 'keyup', (e) => {
      EXPERIMENT.colRange = e.target.value;
      // dom.getById(includeHeadersContainer).style.display = isEntireColumn(e.target.value) ? "flex" : "none";
    });
    dom.addEvent(confidenceDd, 'change', ({ target: { checked } }) => {
      EXPERIMENT.showConfidence = checked;
      // dom.changeText(confidenceValue, checked);
    });
    dom.addEvent(hasHeadersCb, 'change', ({ target: { checked } }) => {
      EXPERIMENT.hasHeaders = checked;
      // dom.changeText(hasHeadersValue, checked);
    });
    dom.addEvent(insertDd, 'change', ({ target: { value } }) => {
      EXPERIMENT.insertPlace = value;
      // dom.changeText(resultPlaceValue, value);
    });
    dom.addEvent(modelDd, 'change', ({ target: { value } }) => {
      findAndSetModelByName(value);
      // dom.changeText(modelValue, value);
    });
    dom.addEvent(runExperiment, 'click', () => {
      if (dom.getById(runExperiment).innerText === 'Stop') {
        dom.changeText(runTotalCells, 0);
        dom.changeText(runProcessed, 0);
        dom.changeText(runPercentage, 0);
        runState.stop = true;
        dom.changeText(runExperiment, 'Run model');
        dom.getById(runEditWarning).style.display = 'none';
        return;
      }
      try {
        runState.clearRunState(dom);
        runExcelAPI(dom);
        dom.getById(runEditWarning).style.display = 'block';
        dom.changeText(runExperiment, 'Stop');
      } catch (e) {
        console.error(e);
        alert(
          'Sorry, something went wrong, please verify you typed a correct range value.\n For example: "A2:A10" of "B:B"'
        );
      }
    });
  };

  setLoginEvents(() => setExperimentEvents(setRunModelPageEvents));

  testToken && setExperimentEvents(setRunModelPageEvents);
  // setRunModelPageEvents();
}

const http = new HttpService(API_URL);
const domHandler = new DomHandler();

// let router = "";
let router = new PageStateHandler(pages.login.pageId, domHandler);
if (testToken) {
  // Uncomment line below to start from experiments page
  router = new PageStateHandler(pages.experiments.pageId, domHandler);

  // Uncomment line below to start from model page
  // router = new PageStateHandler(pages.run.pageId, domHandler);
}

setUpListeners(domHandler, router, pages, http);
