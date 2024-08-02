const { app, BrowserWindow } = require("electron");
const path = require("path");
const { spawn } = require("child_process");
const ngrok = require('ngrok');
const http = require("http");

let ngrokUrl = '';

const createWindow = (url) => {
    const mainWindow = new BrowserWindow({
        width: 1280,
        height: 720,
        webPreferences: {
            nodeIntegration: true,
            contextIsolation: true,
            webSecurity: false,
        },
    });

    mainWindow.loadURL(url).catch((err) => {
        console.error(`Error loading URL: ${url}`, err);
        mainWindow.loadURL('data:text/plain,Failed to load content');
    });

    mainWindow.webContents.on('did-fail-load', (event, errorCode, errorDescription, validatedURL) => {
        console.error(`Failed to load URL: ${validatedURL}, Error: ${errorDescription}`);
        mainWindow.loadURL('data:text/plain,Failed to load content');
    });

    mainWindow.on('closed', () => {
        console.log("Main window closed");
    });

    mainWindow.on('unresponsive', () => {
        console.log("Main window is unresponsive");
    });

    // Handle navigation events
    mainWindow.webContents.on('will-navigate', (event, navigationUrl) => {
        console.debug("Navigation attempt:", navigationUrl);
        if (!navigationUrl.startsWith(ngrokUrl)) {
            console.debug("Navigation attempt blocked:", navigationUrl);
            event.preventDefault();
        }
    });

    mainWindow.webContents.setWindowOpenHandler(({ url }) => {
        console.error("New window creation is blocked.");
        return { action: "deny" };
    });
};

const checkPortAvailability = (port) => {
    return new Promise((resolve) => {
        const server = http.createServer();
        server.listen(port, () => {
            server.close(() => resolve(true));
        });
        server.on('error', () => resolve(false));
    });
};

const startStreamlit = async (initialPort = 8501, maxRetries = 20) => {
    for (let i = 0; i < maxRetries; i++) {
        const port = initialPort + i;
        const available = await checkPortAvailability(port);
        if (available) {
            return new Promise((resolve, reject) => {
                const streamlitFilePath = path.join(__dirname, '..', 'streamlit_app', 'streamlit.py');
                console.log(`Trying to start Streamlit at ${streamlitFilePath} on port ${port}`);

                const streamlit = spawn("streamlit", [
                    "run",
                    streamlitFilePath,
                    "--server.port", port.toString(),
                    "--server.headless", "true",
                    "--server.address", "0.0.0.0"
                ]);

                let started = false;

                streamlit.stdout.on('data', (data) => {
                    const message = data.toString();
                    console.log(`Streamlit output: ${message}`);
                    if (message.includes("You can now view your Streamlit app in your browser")) {
                        started = true;
                        setTimeout(() => resolve(port), 2000); // Wait 2 seconds before resolving
                    }
                });

                streamlit.stderr.on('data', (data) => {
                    const errorMessage = data.toString();
                    console.error(`Streamlit stderr: ${errorMessage}`);
                    if (errorMessage.includes('Port') || errorMessage.includes('already in use')) {
                        streamlit.kill(); // Stop the current attempt
                    } else {
                        // Log error but don't reject immediately to allow retry
                        console.error(`Streamlit stderr: ${errorMessage}`);
                    }
                });

                streamlit.on('error', (error) => {
                    console.error(`Streamlit process error: ${error.message}`);
                    reject(error);
                });

                streamlit.on("close", (code) => {
                    console.log(`Streamlit process exited with code ${code}`);
                    if (!started && (code !== 0 || code === null)) {
                        console.log(`Port ${port} seems to be in use or the process failed. Trying next port...`);
                        resolve(null); // Allow the loop to try the next port
                    }
                });

                setTimeout(() => {
                    if (!started) {
                        console.error(`Streamlit process timed out on port ${port}`);
                        streamlit.kill(); // Stop the current attempt
                        resolve(null); // Allow the loop to try the next port
                    }
                }, 60000);
            });
        } else {
            console.warn(`Port ${port} is not available. Trying next port...`);
        }
    }
    throw new Error('No available ports found after checking multiple options.');
};

app.whenReady().then(async () => {
    try {
        let port = null;
        let retries = 0;
        let maxRetries = 15;
        let success = false;

        while (retries < maxRetries && !success) {
            try {
                port = await startStreamlit(8501 + retries);
                if (port !== null) {
                    success = true; // Successfully started Streamlit
                }
            } catch (error) {
                console.error(`Attempt ${retries + 1} failed:`, error);
            }
            retries++;
        }

        if (!success) {
            throw new Error('All ports are in use or Streamlit failed to start.');
        }

        const localUrl = `http://localhost:${port}/`;
        console.log(`Streamlit is available at ${localUrl}`);

        // Start ngrok and get the public URL
        ngrokUrl = await ngrok.connect({
            proto: 'http',
            addr: port, // The port where Streamlit is running
            authtoken: '2k5ku3aCibDZjyw3Wqa9N0Uc9Io_5Eh5hCsnGGoKvz4KqGuJx', // Optional, if you have an ngrok authtoken
        });

        console.log(`ngrok URL: ${ngrokUrl}`);

        // Load ngrok URL in Electron
        createWindow(ngrokUrl);

    } catch (error) {
        console.error('Failed to start Streamlit or ngrok:', error);
        createWindow('data:text/plain,Failed to start Streamlit or ngrok. Please ensure they are installed and the script path is correct.');
    }

    app.on("activate", () => {
        if (BrowserWindow.getAllWindows().length === 0) {
            createWindow(`http://localhost:8501/`);
        }
    });
});

app.on("window-all-closed", () => {
    if (process.platform !== "darwin") {
        console.log("All windows closed, quitting app.");
        app.quit();
    }
});

app.on("uncaughtException", (error) => {
    console.error("Uncaught Exception:", error);
});

app.on("web-contents-created", (event, contents) => {
    contents.on("will-attach-webview", (event, webPreferences, params) => {
        event.preventDefault();
    });

    contents.on("will-navigate", (event, navigationUrl) => {
        console.debug("Navigation attempt:", navigationUrl);
        if (!navigationUrl.startsWith(ngrokUrl)) {
            console.debug("Navigation attempt blocked:", navigationUrl);
            event.preventDefault();
        }
    });

    contents.setWindowOpenHandler(({ url }) => {
        console.error("New window creation is blocked.");
        return { action: "deny" };
    });
});
