import { defineConfig } from 'vitest/config';

export default defineConfig({
	test: {
		globals: true,
		include: ['tests/**/*.test.ts'],
		coverage: {
			provider: 'v8',
			include: ['src/**/*.ts'],
			exclude: [
				'src/**/*.d.ts',
				'src/index.ts',
				'src/*/index.ts',       // Barrel exports
				'src/xml/preserve.ts',  // Unused utility
				'src/xml/serialize.ts', // Unused utility
			],
			reporter: ['text', 'json', 'html'],
			thresholds: {
				lines: 80,
				functions: 80,
				branches: 80,
				statements: 80,
			},
		},
	},
});
