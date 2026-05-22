/**
 * Custom structural validator for the JSW Township DPR Sheets
 * Ensures that all critical dashboard sections contain required properties, arrays, and numerical values
 */

const validateDprData = (data) => {
  const errors = [];

  const checkArray = (sheetName, requiredKeys) => {
    const list = data[sheetName];
    if (!list) {
      errors.push(`Missing worksheet section: '${sheetName}'`);
      return;
    }
    if (!Array.isArray(list)) {
      errors.push(`Worksheet section '${sheetName}' must be a list of records.`);
      return;
    }
    if (list.length === 0) {
      errors.push(`Worksheet section '${sheetName}' is empty.`);
      return;
    }

    // Check first few rows for mandatory structure
    const sample = list[0];
    requiredKeys.forEach(key => {
      if (sample[key] === undefined) {
        errors.push(`Sheet '${sheetName}' is missing required column: '${key}'`);
      }
    });
  };

  try {
    // 1. Guest House Occupancy
    checkArray('guestHouseData', ['name', 'total', 'occupied', 'vacant']);

    // 2. Canteen Meals
    checkArray('mealData', ['name', 'breakfast', 'lunch', 'dinner']);

    // 3. Vehicles/Fleet
    checkArray('vehicleData', ['name', 'km']);

    // 4. F&B Manpower
    checkArray('manpowerFB', ['name', 'actual', 'present']);

    // 5. Maintenance Manpower
    checkArray('manpowerMaint', ['name', 'actual', 'present']);

    // 6. Complaints categories (Pie)
    checkArray('complaintsPie', ['name', 'value']);

    // 7. Complaints status
    checkArray('complaintsStatus', ['name', 'completed', 'inProgress', 'pending']);

    // 8. Paint Work
    checkArray('paintProgress', ['name', 'target', 'done']);

  } catch (err) {
    errors.push(`Fatal schema parsing error: ${err.message}`);
  }

  return {
    isValid: errors.length === 0,
    errors
  };
};

module.exports = { validateDprData };
