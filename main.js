$("#run").click(() => tryCatch(run));

async function run() {
  async function getUserData() {
    try {
      let userTokenEncoded = await Office.auth.getAccessToken();
      console.log(userTokenEncoded);

      //console.log(userToken.name); // user name
      //console.log(userToken.preferred_username); // email
      //console.log(userToken.oid); // user id
    } catch (exception) {
      if (exception.code === 13003) {
        // SSO is not supported for domain user accounts, only
        // Microsoft 365 Education or work account, or a Microsoft account.
      } else {
        // Handle error
      }
    }
  }
}

async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}
