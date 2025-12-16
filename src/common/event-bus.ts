export class EventBus {
    private eventBus: Map<string, ((data?: any) => Promise<void>)[]> = new Map();

    on(event: string, listener: (data?: any) => Promise<void>) {
        if (!this.eventBus.has(event)) {
            this.eventBus.set(event, []);
        }
        this.eventBus.get(event)?.push(listener);
    }

    off(event: string) {
        this.eventBus.delete(event);
    }

    async emit(event: string, data?: any) {
        const listeners = this.eventBus.get(event);
        if (listeners) {
            for (const listener of listeners) {
                await listener(data);
            }
        }
    }

}