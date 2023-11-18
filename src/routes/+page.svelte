<script>
	import ExcelJS from 'exceljs';
	import moment from 'moment';

	let finishedData = null;
	let loading = false;
	let dates = [];
	let selectedDate = false;

	let reportDate = '';
	let reportUser = '';

	function startLoad(event) {
		loading = true;
		setTimeout(() => {
			processFile(event);
		}, 1000);
	}

	function convertToMilitaryTime(timeString) {
		// Parse the input time string
		const parsedTime = /(\d{1,2}):(\d{2})([APMapm]{2})/.exec(timeString);

		if (parsedTime) {
			let hours = parseInt(parsedTime[1], 10);
			const minutes = parsedTime[2];
			const period = parsedTime[3].toUpperCase();

			// Adjust hours based on AM/PM
			if (period === 'PM' && hours !== 12) {
				hours += 12;
			} else if (period === 'AM' && hours === 12) {
				hours = 0;
			}

			// Format hours and minutes as two digits
			const formattedHours = hours.toString().padStart(2, '0');
			return `${formattedHours}:${minutes}`;
		}

		// Return the input string if it doesn't match the expected format
		return timeString;
	}

	function week(jsonData) {
		for (const key in jsonData) {
			const element = jsonData[key];

			// Check if the properties match the conditions
			if (
				element['1'] === 'Name' &&
				element['2'] === 'Name' &&
				element['3'] === 'Name' &&
				element['5'] === 'Hours' &&
				element['6'] === 'Hours' &&
				element['7'] === 'Hours'
			) {
				return element;
			}
		}

		// Return null if no matching element is found
		return null;
	}

	function processFile(event) {
		const file = event.target.files[0];

		if (file) {
			const reader = new FileReader();

			reader.onload = async function (event) {
				const arrayBuffer = event.target.result;

				// Use exceljs to read the Excel file
				const workbook = new ExcelJS.Workbook();
				await workbook.xlsx.load(arrayBuffer);

				const worksheet = workbook.worksheets[0];
				let jsonData = [];

				worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
					const rowJson = {};
					row.eachCell((cell, colNumber) => {
						rowJson[colNumber] = cell.value;
					});
					jsonData.push(rowJson);
				});

				reportUser = jsonData[2]['19'];
				dates = [
					week(jsonData)['8'].richText[0].text,
					week(jsonData)['11'].richText[0].text,
					week(jsonData)['13'].richText[0].text,
					week(jsonData)['15'].richText[0].text,
					week(jsonData)['16'].richText[0].text,
					week(jsonData)['19'].richText[0].text,
					week(jsonData)['20'].richText[0].text
				];
				Object.keys(jsonData).forEach((key) => {
					const e = jsonData[key];

					if (
						e['1'] === 'Name' &&
						e['2'] === 'Name' &&
						e['3'] === 'Name' &&
						e['5'] === 'Hours' &&
						e['6'] === 'Hours' &&
						e['7'] === 'Hours'
					) {
						delete jsonData[key];
					}
				});
				Object.keys(jsonData).forEach((key) => {
					const e = jsonData[key];

					if (
						e['1'] == 'Time Period :' ||
						e['1'] == 'Query :' ||
						e['1'] == 'Currency Code :' ||
						e['1'] == 'Group Forecast' ||
						e['1'] == 'Fcst.' ||
						e['1'] == 'Sched.' ||
						e['1'] == 'O/U hours' ||
						e['1'] == 'SvF%' ||
						e['10'] == null ||
						e['4'] == 'Error' ||
						(e['1'] == '' && e['5'] > 0) ||
						e['5'] == 0
					) {
						delete jsonData[key];
					}
				});

				let dept = '';
				let shifts = [];
				Object.keys(jsonData).forEach((key) => {
					const e = jsonData[key];
					if (typeof e['8'] === 'string' && e['8'].startsWith('Store')) {
						dept = e['8'].split('Dept:')[1];
					} else {
						const emp = e['1'].richText[0].text;
						const dayStrings = ['8', '11', '13', '15', '16', '19', '20'];
						for (var i = 0; i < 7; i++) {
							if (e[dayStrings[i]] != '') {
								let shift = false;
								let lunch = false;
								let task = false;
								let job = false;
								e[dayStrings[i]].richText.forEach((seg) => {
									if (seg.font.color.argb != 'FFC0C0C0') {
										if (
											(seg.font.size == 5 && /^[0-9]/.test(seg.text)) ||
											seg.text.startsWith('Hrs')
										) {
											if (shift) {
												const shiftArray = {
													employee: emp,
													department: dept,
													shift: shift,
													lunch: lunch,
													task: task,
													job: job,
													day: dates[i],
													start: convertToMilitaryTime(shift.split('-')[0])
												};
												shifts = [...shifts, shiftArray];
												shift = false;
												lunch = false;
												task = false;
												job = false;
											}
											if (!seg.text.startsWith('Hrs')) {
												shift = seg.text;
											}
										} else {
											if (
												!seg.text.startsWith('\\') &&
												!seg.text.startsWith('(') &&
												isNaN(seg.text.charAt(0))
											) {
												if (seg.font.color.argb == 'FF000000') {
													job = seg.text;
												} else task = seg.text;
											} else if (seg.text.startsWith('(')) {
												lunch = seg.text.split('(M) ')[1];
											}
										}
									}
								});
							}
						}
					}
				});
				const sortedData = shifts.sort((a, b) => {
					// Compare by start time
					const startComparison = a.start.localeCompare(b.start);
					if (startComparison !== 0) {
						return startComparison;
					}

					// If start times are the same, compare by employee name
					return a.employee.localeCompare(b.employee);
				});
				// Initialize the nested object
				const nestedDataObject = {};

				// Iterate through the sorted array and build the nested structure
				sortedData.forEach((item) => {
					const { day, department, job } = item;

					// Create or update the nested structure
					nestedDataObject[day] = nestedDataObject[day] || {};
					nestedDataObject[day][department] = nestedDataObject[day][department] || {};
					nestedDataObject[day][department][job] = nestedDataObject[day][department][job] || [];
					nestedDataObject[day][department][job].push(item);
				});

				const today = new Date();
				let dayOfWeek = today.getDay() - 1; // 0 for Monday, 6 for Sunday

				// Adjust for Sunday (which should be 6)
				dayOfWeek = dayOfWeek === -1 ? 6 : dayOfWeek;

				selectedDate = dates[dayOfWeek];
				console.log(sortedData);
				console.log(nestedDataObject);
				finishedData = nestedDataObject;
				loading = false;
			};

			reader.readAsArrayBuffer(file);
		}
	}
</script>

<div class="text-center">
	{#if loading}
		<div class="font-medium mt-5 text-[25px]">
			Hang tight! Our data hamsters are running as fast as they can...
		</div>
		<div class="hamster text-[200px]">🐹</div>
	{:else if finishedData}
		<div class="flex justify-center items-center screen-only">
			<div class="w-2/5 grid grid-cols-2 gap-3">
				<div
					class="cursor-pointer text-white my-3 bg-blue-700 hover:bg-blue-800 focus:outline-none focus:ring-4 focus:ring-blue-300 font-medium rounded-full text-sm w-50 py-2.5 text-center mb-2 dark:bg-blue-600 dark:hover:bg-blue-700 dark:focus:ring-blue-800"
					on:click={() => window.print()}
				>
					Print Report
				</div>
				<div
					class="cursor-pointer text-white my-3 bg-blue-700 hover:bg-blue-800 focus:outline-none focus:ring-4 focus:ring-blue-300 font-medium rounded-full text-sm w-50 py-2.5 text-center mb-2 dark:bg-blue-600 dark:hover:bg-blue-700 dark:focus:ring-blue-800"
					on:click={() => window.location.reload()}
				>
					Start Over
				</div>
			</div>
		</div>
		<div class="screen-only my-3">
			<select
				class="h-[35px] mt-auto border-2 border-blue-500"
				style="font-size:25px !important;"
				bind:value={selectedDate}
				name="sel"
				id="sel"
			>
				{#each dates as item}
					<option value={item}>{item}</option>
				{/each}
			</select>
		</div>
		<div class="header mx-[20%] print:mx-0">
			<div class="grid grid-cols-3 text-[14px] mb-3">
				<div class="text-left">Date: {selectedDate}</div>
				<div class="font-bold">Store Coverage by Day</div>
				<div class="text-right">Generated By: {reportUser}</div>
			</div>
		</div>
		<div class="text-[18px] print:text-[10px] mx-[20%] print:mx-0" id="display-area">
			<div class="text-left">
				{#each Object.keys(finishedData[selectedDate]).sort() as department}
					<div class="column-break mb-15px">
						<div class="border border-black bg-gray-800 text-white text-center font-bold">
							{department}
						</div>
						{#each Object.keys(finishedData[selectedDate][department]) as job}
							{#if job != 'Associate'}
								<div
									class="border-l border-r border-b border-black bg-gray-400 text-center font-bold"
								>
									{job}
								</div>
							{/if}
							{#each finishedData[selectedDate][department][job] as shift}
								<div class="flex flex-row border-b border-black overflow-hidden whitespace-nowrap">
									<div
										class="overflow-hidden whitespace-nowrap px-1 w-[20%] border-l border-black uppercase"
									>
										{shift.employee}
									</div>
									<div
										class="overflow-hidden whitespace-nowrap px-1 w-[30%] border-l border-black uppercase"
									>
										{moment(
											'2023-12-12 ' + convertToMilitaryTime(shift.shift.split('-')[0])
										).format('hh:mm A') +
											' - ' +
											moment(
												'2023-12-12 ' + convertToMilitaryTime(shift.shift.split('-')[1])
											).format('hh:mm A')}
									</div>
									<div
										class="overflow-hidden whitespace-nowrap px-1 w-[30%] border-l border-black uppercase"
									>
										{shift.lunch
											? moment(
													'2023-12-12 ' + convertToMilitaryTime(shift.lunch.split('-')[0])
											  ).format('hh:mm A') +
											  ' - ' +
											  moment(
													'2023-12-12 ' + convertToMilitaryTime(shift.lunch.split('-')[1])
											  ).format('hh:mm A')
											: ''}
									</div>
									<div
										class="overflow-hidden whitespace-nowrap px-1 w-[20%] border-l border-black border-r uppercase"
									>
										{shift.task == false ? '' : shift.task}
									</div>
								</div>
							{/each}
						{/each}
					</div>
				{/each}
			</div>
		</div>
	{:else}
		<div class="mb-2 mt-4 flex justify-center">
			<img
				width="90"
				height="90"
				src="https://corporate.homedepot.com/sites/default/files/image_gallery/THD_logo.jpg"
				alt=""
			/>
		</div>
		<div class="text-[20pt] mb-3 font-bold text-orange-500">
			Store Coverage by Day Report Generator
		</div>
		<div class="flex justify-center">
			<div class="w-2/5 bg-orange-200 py-3">
				<input type="file" id="excel-file" accept=".xlsx, .xls" on:change={startLoad} />
			</div>
		</div>

		<div class="text-xl text-orange-400 mt-3">
			Upload .xlsx Dimensions Store Coverage By Day above to continue...
		</div>
	{/if}
</div>

<style>
	* {
		font-family: 'PT Sans Narrow', sans-serif;
	}
	.hamster {
		display: inline-block;
		animation: bounce 1s ease-in-out infinite;
	}

	@keyframes bounce {
		0%,
		100% {
			transform: translateY(0);
		}
		50% {
			transform: translateY(-20px);
		}
	}
	@media print {
		.page-break {
			page-break-after: always;
		}
	}
	@media screen {
		.print-only {
			display: none;
		}
		* {
			font-size: 18px !important;
		}
		.hamster {
			font-size: 200px !important;
		}
	}
	@media print {
		.screen-only {
			display: none;
		}
		* {
			font-size: 9px !important;
		}
	}
	@media print {
		#display-area {
			columns: 2; /* Set the number of columns for printing */
		}
		.column-break {
			break-inside: avoid-column; /* Avoid widows by starting a new group on a new column */
		}
	}
</style>
