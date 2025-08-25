/**
 * フラグ1の値を変更して保存する関数
 */
function setFlag1(value) {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty('key1', value); // キー名を 'key1' に変更
}

/**
 * フラグ2の値を変更して保存する関数
 */
function setFlag2(value) {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty('key2', value); // キー名を 'key2' に変更
}

/**
 * フラグ1の現在の値を取得する関数
 */
function getFlag1() {
  const properties = PropertiesService.getScriptProperties();
  const flagValue = properties.getProperty('key1');
  Logger.log('フラグ1の値: ' + flagValue);
  return flagValue;
}

/**
 * フラグ2の現在の値を取得する関数
 */
function getFlag2() {
  const properties = PropertiesService.getScriptProperties();
  const flagValue = properties.getProperty('key2');
  Logger.log('フラグ2の値: ' + flagValue);
  return flagValue;
}

/** 
const a = getFlag1();
const b = getFlag2();
console.log(a);
console.log(b);
*/