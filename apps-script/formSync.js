function refreshFormLists_() {
  try {
    const form = FormApp.openById(CONFIG.FORM_ID);

    // --- Update Infractions ---
    const rubric = sh_(CONFIG.TABS.RUBRIC);
    const vals = rubric.getRange(2, 1, rubric.getLastRow() - 1, 4).getValues();
    // Columns: [Infraction, Points, RedLine, Notes/PolicyLink]

    let infractions = vals.map(r => {
      const infraction = String(r[0] || '').trim();
      const points     = String(r[1] || '').trim();
      const redline    = String(r[2] || '').trim();
      const policy     = String(r[3] || '').trim();

      if (!infraction) return null;

      // Build label
      let label = '';
      if (policy) label += `[${policy}] `;
      label += infraction;
      if (points) label += ` — ${points} pts`;

      return label;
    }).filter(v => v);

    // Deduplicate
    infractions = [...new Set(infractions)];

    const infractionItem = form.getItems().find(it => it.getTitle() === 'Infraction');
    if (infractionItem) {
      infractionItem.asListItem().setChoiceValues(infractions);
    }

    // --- Update Employees & Managers ---
    const dv = sh_(CONFIG.TABS.DATA_VALIDATION);
    let employees = dv.getRange(2, 1, dv.getLastRow() - 1, 1).getValues()
                      .map(r => String(r[0]).trim()).filter(v => v);
    employees = [...new Set(employees)];

    let managers = dv.getRange(2, 2, dv.getLastRow() - 1, 1).getValues()
                     .map(r => String(r[0]).trim()).filter(v => v);
    managers = [...new Set(managers)];

    const empItem = form.getItems().find(it => it.getTitle() === 'Employee');
    if (empItem) empItem.asListItem().setChoiceValues(employees);

    const mgrItem = form.getItems().find(it => it.getTitle() === 'Lead');
    if (mgrItem) mgrItem.asListItem().setChoiceValues(managers);

    logAudit('system', 'refreshFormLists_', null, {
      infractions: infractions.length,
      employees: employees.length,
      managers: managers.length
    });

    SpreadsheetApp.getUi().alert(
      `✅ Form updated!\nInfractions: ${infractions.length}\nEmployees: ${employees.length}\nLeads: ${managers.length}`
    );

  } catch (e) {
    logError('refreshFormLists_', e);
  }
}
