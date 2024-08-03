/**
 * Copyright 2024 k.k.Factory
 */
export const isNullish = (value: unknown): value is null | undefined => {
	return value === null || value === undefined;
};
