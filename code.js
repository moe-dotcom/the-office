// jsfiddle draft https://jsfiddle.net/1x4rhL2n/

document.addEventListener('DOMContentLoaded', function() {
    const dropArea = document.getElementById('drop-area');
    const fileElem = document.getElementById('fileElem');
    const fileList = document.getElementById('file-list');
    const analyseBtn = document.getElementById('analyse-btn');
    const errorMessage = document.getElementById('error-message');
    
    let files = [];
    
    // Prevent default drag behaviors
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, preventDefaults, false);
    });
    
    // Highlight drop area when item is dragged over it
    ['dragenter', 'dragover'].forEach(eventName => {
        dropArea.addEventListener(eventName, highlight, false);
    });
    
    ['dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, unhighlight, false);
    });
    
    // Handle dropped files
    dropArea.addEventListener('drop', handleDrop, false);
    
    // Handle clicked files
    dropArea.addEventListener('click', () => fileElem.click());
    fileElem.addEventListener('change', handleFiles);
    
    // Analyse button click
    analyseBtn.addEventListener('click', analyseFiles);
    
    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }
    
    function highlight() {
        dropArea.classList.add('highlight');
    }
    
    function unhighlight() {
        dropArea.classList.remove('highlight');
    }
    
    function handleDrop(e) {
        const dt = e.dataTransfer;
        handleFiles({ target: { files: dt.files } });
    }
    
    function handleFiles(e) {
        files = [...files, ...Array.from(e.target.files)];
        displayFileList();
        analyseBtn.disabled = files.length < 2;
    }
    
    function displayFileList() {
        fileList.innerHTML = '<h3>Selected Files:</h3><ul>' + 
            files.map(file => `<li>${file.name}</li>`).join('') + 
            '</ul>';
    }
    
    async function analyseFiles() {
        errorMessage.textContent = '';
        document.getElementById('swiped-not-checked-in').innerHTML = '';
        document.getElementById('swiped-no-booking').innerHTML = '';
        
        try {
            // Find booking and swipe files
            const bookingFile = files.find(file => 
                file.name.toLowerCase().includes('booking') || 
                file.name.toLowerCase().includes('bookings')
            );
            
            const swipeFile = files.find(file => 
                file.name.toLowerCase().includes('swipe') ||
                file.name.toLowerCase().includes('entry') ||  
                file.name.toLowerCase().includes('access')
            );
            
            if (!bookingFile || !swipeFile) {
                throw new Error('Please include both a bookings file and a swipe data file');
            }
            
            // Read files
            const bookingData = await readExcel(bookingFile);
            const swipeData = await readExcel(swipeFile);
            
            // Process data
            const bookings = processBookingData(bookingData);
            const swipes = processSwipeData(swipeData);
            
            // Analyse
            const results = analyseData(bookings, swipes);
            
            // Display results
            displayResults(results);
            
        } catch (error) {
            errorMessage.textContent = 'Error: ' + error.message;
            console.error(error);
        }
    }
    
    function readExcel(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    
                    // Get first sheet
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonData = XLSX.utils.sheet_to_json(firstSheet);
                    
                    resolve(jsonData);
                } catch (error) {
                    reject(error);
                }
            };
            
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    }
    
    function processBookingData(data) {
        // Normalize booking data - adjust column names as needed
        return data.map(row => {
            // Try to find name and status columns (case insensitive)
            const nameKey = Object.keys(row).find(key => 
                key.toLowerCase().includes('name') || 
                key.toLowerCase().includes('employee')
            );
            
            const statusKey = Object.keys(row).find(key => 
                key.toLowerCase().includes('status') || 
                key.toLowerCase().includes('check') ||
                key.toLowerCase().includes('attended')
            );
            
            if (!nameKey || !statusKey) {
                throw new Error('Booking file must contain name and status columns');
            }
            
            return {
                name: String(row[nameKey]).trim(),
                checkedIn: String(row[statusKey]).toLowerCase().includes('check') || 
                          String(row[statusKey]).toLowerCase().includes('yes') ||
                          String(row[statusKey]).toLowerCase() === 'true'
            };
        });
    }
    
    function processSwipeData(data) {
    // Object to store unique swipes (first swipe only per person)
        const uniqueSwipes = {};
    
        data.forEach(row => {
            const nameKey = Object.keys(row).find(key => 
                key.toLowerCase().includes('name') || 
                key.toLowerCase().includes('text')
        );
        
            const timeKey = Object.keys(row).find(key => 
                key.toLowerCase().includes('time') || 
                key.toLowerCase().includes('swipe') ||
                key.toLowerCase().includes('access')
        );
        
            if (!nameKey || !timeKey) {
                throw new Error('Swipe file must contain name and time columns');
        }
        
            const name = String(row[nameKey]).trim();
            const time = row[timeKey];
        
        // Skip if we already have a swipe for this person
            if (!uniqueSwipes[name]) {
                uniqueSwipes[name] = {
                    name: name,
                    time: convertExcelTime(time) // Convert Excel time to proper format
                };
            }
    });
    
    // Return just the values (array of swipe objects)
        return Object.values(uniqueSwipes);
}

    // Helper function to convert Excel serial date/time to JS Date
    function convertExcelTime(excelTime) {
        // If it's already a Date object or string, return as-is
        if (excelTime instanceof Date) return excelTime;
        if (typeof excelTime === 'string') return excelTime;
    
    // If it's a number (Excel serial date format)
        if (typeof excelTime === 'number') {
    // Excel dates are based on 1900-01-01 (with a bug for 1900 being a leap year)
            const excelEpoch = new Date(1899, 11, 30);
            const jsDate = new Date(excelEpoch.getTime() + excelTime * 24 * 60 * 60 * 1000);
        
    // Format as readable time (adjust format as needed)
            return jsDate.toLocaleString();
    }
    
    // Fallback for other cases
            return excelTime;
}
    
    function analyseData(bookings, swipes) {
        const results = {
            swipedButNotCheckedIn: [],
            swipedWithoutBooking: []
        };
        
        // Create a map of bookings by name
        const bookingMap = {};
        bookings.forEach(booking => {
            bookingMap[booking.name] = booking;
        });
        
        // Analyse each swipe
        swipes.forEach(swipe => {
            const booking = bookingMap[swipe.name];
            
            if (booking) {
                // Person has a booking
                if (!booking.checkedIn) {
                    results.swipedButNotCheckedIn.push({
                        name: swipe.name,
                        swipeTime: swipe.time,
                        bookingStatus: 'Did not check in'
                    });
                }
            } else {
                // No booking found
                results.swipedWithoutBooking.push({
                    name: swipe.name,
                    swipeTime: swipe.time
                });
            }
        });
        
        return results;
    }
    
    function displayResults(results) {
        // Display people who swiped but didn't check in
        const swipedNotCheckedIn = document.getElementById('swiped-not-checked-in');
        if (results.swipedButNotCheckedIn.length > 0) {
            swipedNotCheckedIn.innerHTML = createTable(
                ['Email', 'Swipe Time', 'Booking Status'],
                results.swipedButNotCheckedIn.map(item => [
                    item.name, 
                    formatTime(item.swipeTime), 
                    item.bookingStatus
                ])
            );
        } else {
            swipedNotCheckedIn.innerHTML = '<p>No people found in this category</p>';
        }
        
        // Display people who swiped without a booking
        const swipedNoBooking = document.getElementById('swiped-no-booking');
        if (results.swipedWithoutBooking.length > 0) {
            swipedNoBooking.innerHTML = createTable(
                ['Name', 'Swipe Time'],
                results.swipedWithoutBooking.map(item => [
                    item.name, 
                    formatTime(item.swipeTime)
                ])
            );
        } else {
            swipedNoBooking.innerHTML = '<p>No people found in this category</p>';
        }
    }
    
    function createTable(headers, rows) {
        return `
            <table>
                <thead>
                    <tr>${headers.map(h => `<th>${h}</th>`).join('')}</tr>
                </thead>
                <tbody>
                    ${rows.map(row => 
                        `<tr>${row.map(cell => `<td>${cell}</td>`).join('')}</tr>`
                    ).join('')}
                </tbody>
            </table>
        `;
    }
    
    function formatTime(time) {
        if (!time) return 'N/A';
        
        // If it's already formatted (from convertExcelTime)
        if (typeof time === 'string') return time;
        
        // If it's a Date object
        if (time instanceof Date) return time.toLocaleString();
        
        // If it's a number that wasn't converted (shouldn't happen after our updates)
        if (typeof time === 'number') {
            return convertExcelTime(time); // Re-use our conversion function
        }
        
        // Fallback - just convert to string
        return String(time);
    }
});