export function filterData(value, data, filterColumn) {
  return value
    ? data.filter(i => i[filterColumn].toLowerCase().indexOf(value) > -1)
    : data;
}

function getNestedObject(nestedObj, pathArr) {
  return pathArr.reduce(
    (obj, key) => (obj && obj[key] !== "undefined" ? obj[key] : undefined),
    nestedObj
  );
}

export function getUniqueValues(data, column) {
  const pathArr = column.split(".");
  let values = data.map(item => {
    return getNestedObject(item, pathArr);
  });
  return values.filter((item, i, ar) => {
    return ar.indexOf(item) === i;
  });
}

export function sortData(columns, sortedCol, items) {
  let isSortedDescending = sortedCol.isSortedDescending;

  // If we've sorted this column, flip it.
  if (sortedCol.isSorted) {
    isSortedDescending = !isSortedDescending;
  }
  const pathArr = sortedCol.fieldName.split(".");
  // Sort the items.
  items = items!.concat([]).sort((a, b) => {
    const firstValue = getNestedObject(a, pathArr);
    const secondValue = getNestedObject(b, pathArr);
    if (isSortedDescending) {
      return firstValue > secondValue ? -1 : 1;
    } else {
      return firstValue > secondValue ? 1 : -1;
    }
  });
  let updatedColumns = columns!.map(col => {
    col.isSorted = col.key === sortedCol.key;
    if (col.isSorted) {
      col.isSortedDescending = isSortedDescending;
    }
    return col;
  });
  return {
    items,
    columns: updatedColumns
  };
}
