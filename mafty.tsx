require('core-js/modules/es.promise');
require('core-js/modules/es.string.includes');
require('core-js/modules/es.object.assign');
require('core-js/modules/es.object.keys');
require('core-js/modules/es.symbol');
require('core-js/modules/es.symbol.async-iterator');
require('regenerator-runtime/runtime');

const Excel = require('exceljs');
const readline = require('readline-sync');

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
	fecha_remuneracion: string;
	fecha_primero: string;
};

type empresaData = {
	nombre: string;
	numero_mtss: string;
	rut: string;
	grupo: string;
	subgrupo: string;
};

type fechaUruguay = {
	dd: string;
	mm: string;
	yyyy: string;
};

enum Files {
	FileToWrite = './reciboALlenar.xlsx',
	WorkersData = './datosTrabajadores.xlsx',
	BusinessData = './datosEmpresa.xlsx',
}

const trabajadoresAProcesar: trabajadorType[] = [];

let empresaDataCeldas: empresaData = {
	nombre: '',
	numero_mtss: '',
	rut: '',
	grupo: '',
	subgrupo: '',
};

let empresaDataAAplicar: empresaData = {
	nombre: '',
	numero_mtss: '',
	rut: '',
	grupo: '',
	subgrupo: '',
};

let celdasAEditar: celdasEditarType = {
	ci: '',
	nombre: '',
	cargo: '',
	fecha_ingreso: '',
	afiliacion_bps: '',
	sueldo_nominal: '',
	fonasa: '',
	fecha_cargo: '',
	fecha_remuneracion: '',
	fecha_primero: '',
};

async function writeOnCell(
	cell: string,
	value: any,
	file?: string,
	newFile?: string,
	workbookP?: any,
) {
	const fileToRead = file ?? Files.FileToWrite;
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

async function getDatosTrabajadores(file = Files.WorkersData) {
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
					fecha_remuneracion: row[9],
					fecha_primero: row[10],
				};
				return;
			}

			//Ignoro la segunda y tercer fila que es info para el usuario
			if (rowNumber == 2) return;
			if (rowNumber == 3) return;

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
	} catch (error) {
		console.error(error);
	}
}

async function crearArchivosParaTrabajadores() {
	for (const trabajador of trabajadoresAProcesar) {
		const fechasArray: string[] = trabajador.fecha_cargo.split('/');

		const fechaCargo = {
			dd: fechasArray[0],
			mm: fechasArray[1],
			yyyy: fechasArray[2],
		};

		const fechaRemuneracion = `${fechaCargo.mm}/${fechaCargo.yyyy}`;

		const fileToWrite = `./ExcelsAImprimir/${trabajador.nombre}--${fechaCargo.dd}-${fechaCargo.mm}-${fechaCargo.yyyy}.xlsx`;

		try {
			const newFileToWrite = new Excel.Workbook();

			newFileToWrite.xlsx.writeFile(fileToWrite);

			await writeOnCell(
				celdasAEditar.ci,
				trabajador.ci,
				undefined,
				fileToWrite,
			);
			await writeOnCell(
				celdasAEditar.nombre,
				trabajador.nombre,
				fileToWrite,
				fileToWrite,
			);
			await writeOnCell(
				celdasAEditar.cargo,
				trabajador.cargo,
				fileToWrite,
				fileToWrite,
			);
			await writeOnCell(
				celdasAEditar.fecha_ingreso,
				trabajador.fecha_ingreso,
				fileToWrite,
				fileToWrite,
			);
			await writeOnCell(
				celdasAEditar.afiliacion_bps,
				trabajador.afiliacion_bps,
				fileToWrite,
				fileToWrite,
			);
			await writeOnCell(
				celdasAEditar.sueldo_nominal,
				trabajador.sueldo_nominal,
				fileToWrite,
				fileToWrite,
			);
			await writeOnCell(
				celdasAEditar.fonasa,
				trabajador.fonasa,
				fileToWrite,
				fileToWrite,
			);
			await writeOnCell(
				celdasAEditar.fecha_cargo,
				trabajador.fecha_cargo,
				fileToWrite,
				fileToWrite,
			);
			await writeOnCell(
				celdasAEditar.fecha_remuneracion,
				fechaRemuneracion,
				fileToWrite,
				fileToWrite,
			);

			await recalcularFormulas(fileToWrite);

			console.log(`Recibo de ${trabajador.nombre} generado!`);
		} catch (error) {
			console.error(error);
			console.log('Algo malo ha ocurrio procesando a ' + trabajador.nombre);
		}
	}
}

async function recalcularFormulas(file: Files | string) {
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

async function actualizarDatosEmpresa(file = Files.BusinessData) {
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

		await writeOnCell(
			empresaDataCeldas.nombre,
			empresaDataAAplicar.nombre,
			Files.FileToWrite,
			Files.FileToWrite,
		);
		await writeOnCell(
			empresaDataCeldas.numero_mtss,
			empresaDataAAplicar.numero_mtss,
			Files.FileToWrite,
			Files.FileToWrite,
		);
		await writeOnCell(
			empresaDataCeldas.rut,
			empresaDataAAplicar.rut,
			Files.FileToWrite,
			Files.FileToWrite,
		);
		await writeOnCell(
			empresaDataCeldas.grupo,
			empresaDataAAplicar.grupo,
			Files.FileToWrite,
			Files.FileToWrite,
		);
		await writeOnCell(
			empresaDataCeldas.subgrupo,
			empresaDataAAplicar.subgrupo,
			Files.FileToWrite,
			Files.FileToWrite,
		);

		await recalcularFormulas(Files.FileToWrite);
	} catch (error) {
		console.error(error);

		console.log('Error al actualizar datos empresa.');
	}
}

async function main() {
	readline.question(
		'Apretar cualquier tecla para iniciar. Salga cerrando esta ventana.',
	);

	console.log('Leyendo archivo de empresa...');
	await actualizarDatosEmpresa();

	console.log('Leyendo archivo de trabajadores...');
	await getDatosTrabajadores();

	console.log('Creando recibos excel de trabajadores...');
	await crearArchivosParaTrabajadores();

	console.log('¡Finalizado!');

	readline.question('Apretar cualquier tecla para salir.');
}

main();
