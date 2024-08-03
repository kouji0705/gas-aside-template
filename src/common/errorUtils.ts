/**
 * Copyright 2024 k.k.Factory
 */
export const createSheetNotFoundError = (sheetName: string): Error => {
	return new Error(`sheetName: ${sheetName} is not found.`);
};

export const createInvalidPositionError = (position: string): Error => {
	return new Error(`position: ${position} is invalid format.`);
};

export const createInvalidPositionsError = (positions: string): Error => {
	return new Error(`positions: ${positions} is invalid format.`);
};
