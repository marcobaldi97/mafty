require("core-js/modules/es.promise");
require("core-js/modules/es.string.includes");
require("core-js/modules/es.object.assign");
require("core-js/modules/es.object.keys");
require("core-js/modules/es.symbol");
require("core-js/modules/es.symbol.async-iterator");
require("regenerator-runtime/runtime");

type trabajadorType = {
	ci: any;
	nombre: any;
	cargo: any;
	fecha_ingreso: any;
	afiliacion_bps: any;
	sueldo_nominal: any;
	fonasa: any;
	fecha_cargo: any;
};

type celdasEditarType = {
	ci: string;
	nombre: string;
	cargo: string;
	fecha_ingreso: string;
	afiliacion_bps: string;
	sueldo_nominal: string;
	fonasa: string;
	fecha_cargo: string;
};

type empresaData = {
	nombre: string;
	numero_mtss: string;
	rut: string;
	grupo: string;
	subgrupo: string;
};

const Excel = require("exceljs");
const defaultFileToWrite = "./reciboALlenar.xlsx";
const defaultFileToRead = "./datosTrabajadores.xlsx";
const defaultBusinessData = "./datosEmpresa.xlsx";
const trabajadoresAProcesar: trabajadorType[] = [];

let empresaDataCeldas: empresaData = {
	nombre: "",
	numero_mtss: "",
	rut: "",
	grupo: "",
	subgrupo: "",
};

let empresaDataAAplicar: empresaData = {
	nombre: "",
	numero_mtss: "",
	rut: "",
	grupo: "",
	subgrupo: "",
};

let celdasAEditar: celdasEditarType = {
	ci: "",
	nombre: "",
	cargo: "",
	fecha_ingreso: "",
	afiliacion_bps: "",
	sueldo_nominal: "",
	fonasa: "",
	fecha_cargo: "",
};

async function writeOnCell(cell: string, value: any, file?: string, newFile?: string, workbookP?: any) {
	const fileToRead = file ?? defaultFileToWrite;
	try {
		const workbook = new Excel.Workbook();
		await workbook.xlsx.readFile(fileToRead);
		const worksheet = workbook.getWorksheet();
		worksheet.getCell(cell).value = value;
		const fileToWrite = newFile ?? file;
		await workbook.xlsx.writeFile(fileToWrite);
	} catch (e) {
		console.error(e);

		console.log(`Error!`);
	}
}

async function getDatosTrabajadores(file?: string) {
	file = file ?? defaultFileToRead;
	try {
		const workbook = new Excel.Workbook();

		await workbook.xlsx.readFile(file);

		workbook.getWorksheet().eachRow(function (row: any, rowNumber: any) {
			row = row.values;
			if (rowNumber == 1) {
				celdasAEditar = {
					ci: row[1],
					nombre: row[2],
					cargo: row[3],
					fecha_ingreso: row[4],
					afiliacion_bps: row[5],
					sueldo_nominal: row[6],
					fonasa: row[7],
					fecha_cargo: row[8],
				};
				return;
			}
			if (rowNumber == 2) return;

			const trabajadorToPush: trabajadorType = {
				ci: row[1],
				nombre: row[2],
				cargo: row[3],
				fecha_ingreso: row[4],
				afiliacion_bps: row[5],
				sueldo_nominal: row[6],
				fonasa: row[7],
				fecha_cargo: row[8],
			};

			trabajadoresAProcesar.push(trabajadorToPush);
		});
	} catch (e) {
		console.log(`Error!`);
		console.log(e);
	}
}

async function crearArchivosParaTrabajadores() {
	for (let index = 0; index < trabajadoresAProcesar.length; index++) {
		const trabajador = trabajadoresAProcesar[index];
		const dateNow = new Date();
		const fechaRemuneracion = `${dateNow.getMonth()}/${dateNow.getFullYear()}`;

		const fileToWrite = `./ExcelsAImprimir/${trabajador.nombre}--${dateNow.getDay()}-${dateNow.getMonth()}-${dateNow.getFullYear()}.xlsx`;
		try {
			const newFileToWrite = new Excel.Workbook();

			newFileToWrite.xlsx.writeFile(fileToWrite);

			await writeOnCell(celdasAEditar.ci, trabajador.ci, undefined, fileToWrite);
			await writeOnCell(celdasAEditar.nombre, trabajador.nombre, fileToWrite, fileToWrite);
			await writeOnCell(celdasAEditar.cargo, trabajador.cargo, fileToWrite, fileToWrite);
			await writeOnCell(celdasAEditar.fecha_ingreso, trabajador.fecha_ingreso, fileToWrite, fileToWrite);
			await writeOnCell(celdasAEditar.afiliacion_bps, trabajador.afiliacion_bps, fileToWrite, fileToWrite);
			await writeOnCell(celdasAEditar.sueldo_nominal, trabajador.sueldo_nominal, fileToWrite, fileToWrite);
			await writeOnCell(celdasAEditar.fonasa, trabajador.fonasa, fileToWrite, fileToWrite);
			await writeOnCell(celdasAEditar.fecha_cargo, trabajador.fecha_cargo, fileToWrite, fileToWrite);
			await writeOnCell("E2", fechaRemuneracion, fileToWrite, fileToWrite);

			await recalcularFormulas(fileToWrite);
			console.log(`Recibo de ${trabajador.nombre} procesado!`);
		} catch (error) {
			console.log(error);
			console.log("Algo malo ocurrio procesando a " + trabajador.nombre);
		}
	}
}

async function recalcularFormulas(file: any) {
	try {
		const workbook = new Excel.Workbook();

		await workbook.xlsx.readFile(file);

		const worksheet = workbook.getWorksheet();

		worksheet.eachRow((row: any) => {
			row.eachCell((cell: any) => {
				if (cell.model.result !== undefined) cell.model.result = undefined;
			});
		});

		await workbook.xlsx.writeFile(file);
	} catch (error) {
		console.error(error);
	}
}

async function actualizarDatosEmpresa(file?: any) {
	file = file ?? defaultBusinessData;

	try {
		let workbook = new Excel.Workbook();

		await workbook.xlsx.readFile(file);

		workbook.getWorksheet().eachRow(function (row: any, rowNumber: any) {
			row = row.values;

			switch (rowNumber) {
				case 1:
					empresaDataCeldas = {
						nombre: row[1],
						numero_mtss: row[2],
						rut: row[3],
						grupo: row[4],
						subgrupo: row[5],
					};
					break;
				case 2:
					break;
				default:
					empresaDataAAplicar = {
						nombre: row[1],
						numero_mtss: row[2],
						rut: row[3],
						grupo: row[4],
						subgrupo: row[5],
					};
					break;
			}
		});

		await writeOnCell(empresaDataCeldas.nombre, empresaDataAAplicar.nombre, defaultFileToWrite, defaultFileToWrite);
		await writeOnCell(empresaDataCeldas.numero_mtss, empresaDataAAplicar.numero_mtss, defaultFileToWrite, defaultFileToWrite);
		await writeOnCell(empresaDataCeldas.rut, empresaDataAAplicar.rut, defaultFileToWrite, defaultFileToWrite);
		await writeOnCell(empresaDataCeldas.grupo, empresaDataAAplicar.grupo, defaultFileToWrite, defaultFileToWrite);
		await writeOnCell(empresaDataCeldas.subgrupo, empresaDataAAplicar.subgrupo, defaultFileToWrite, defaultFileToWrite);

		await recalcularFormulas(defaultFileToWrite);
	} catch (error) {
		console.error(error);

		console.log("Error al actualizar datos empresa.");
	}
}

async function main() {
	console.log("Leyendo archivo de empresa...");
	await actualizarDatosEmpresa();
	console.log("Leyendo archivo de trabajadores...");
	await getDatosTrabajadores();
	console.log("Creando recibos excel de trabajadores...");
	await crearArchivosParaTrabajadores();
	console.log("¡Finalizado!");
}

main();
