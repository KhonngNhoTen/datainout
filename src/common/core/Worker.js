module.exports =
  /** @param {Promise<any>} handler*/
  async function (handler, ...params) {
    await handler(...params);
  };
