window.addEventListener('load', () => {
	console.log('Website loaded');
	function showToast(message, err) {
		const toastContainer = document.querySelector('body');
		// Create a new toast element
		const toast = document.createElement('div');
		toast.textContent = message;
		// Apply CSS styles to the toast element
		if (err) {
			toast.style.backgroundColor = '#C70039';
		} else {
			toast.style.backgroundColor = '#22CB5C';
		}
		toast.style.color = '#fff';
		toast.style.padding = '12px 20px';
		toast.style.borderRadius = '4px';
		toast.style.marginBottom = '10px';
		toast.style.opacity = '0';
		toast.style.transition = 'opacity 0.3s ease';
		toast.style.zIndex = '10000';
		toast.style.position = 'absolute';
		toast.style.top = '10px';
		toast.style.left = '10px';
		// Show the toast
		setTimeout(() => {
			toast.style.opacity = '1';
		}, 100);
		// Append the toast to the container
		toastContainer.appendChild(toast);
		// Remove the toast after 3 seconds
		setTimeout(() => {
			toast.style.opacity = '0';
			// Remove the toast element from the DOM after the transition ends
			toast.addEventListener('transitionend', () => {
				toast.remove();
			});
		}, 5000);
	}
	// Usage:
	async function postData(data) {
		console.log('Posting data ' + JSON.stringify(data));
		const result = await fetch('http://localhost:5000/storeData', {
			body: JSON.stringify(data),
			method: 'POST',
			headers: {
				'Content-Type': 'application/json',
				'Access-Control-Allow-Origin': '*',
			},
		})
			.then((res) => res.json())
			.then((result) => {
				showToast(result?.message || 'Success!');
				return result;
			})
			.catch((err) => {
				console.log('Error from backend ' + err);
				showToast(err?.message || 'Error!', err);
			});
		console.log(result);
	}
	function copyToClipboard(text) {
		navigator.clipboard
			.writeText(text)
			.then(() => {
				console.log('Text copied to clipboard');
			})
			.catch((error) => {
				console.error('Failed to copy text to clipboard:', error);
			});
	}
	const targetElement = document.querySelector('body');
	const observer = new MutationObserver(function (mutationsList, observer) {
		for (let mutation of mutationsList) {
			if (document.querySelector('.hj')?.lastChild?.tagName !== 'BUTTON') {
				const respondedBtn = document.createElement('button');
				respondedBtn.innerText = 'Responded';
				const container = document.querySelector('.hj');
				// container.style.display = 'flex';
				// container.style.alignItems = 'center';

				let email = document
					.getElementsByClassName('go')[0]
					?.innerText.substring(
						1,
						document.getElementsByClassName('go')[0].innerText.length - 1,
					);
				if (!email) {
					const showDetailBtn = document.querySelector(
						"[aria-label='Show details']",
					);
					if (showDetailBtn) {
						showDetailBtn.click();
						showDetailBtn.click();
					}
					email = document
						.getElementsByClassName('go')[0]
						?.innerText.substring(
							1,
							document.getElementsByClassName('go')[0].innerText.length - 1,
						);
				}
				// Getting the email address of Forwarded message from the email content
				if (
					email &&
					(email?.includes('@ineuron.ai') ||
						email?.includes('@pwskills.com') ||
						email?.includes('@pw.live'))
				) {
					let emailContent;
					if (
						document
							.querySelector('.iA.g6')
							?.innerText?.includes('Forwarded message')
					) {
						emailContent = document.querySelector('.iA.g6')?.innerText;
					} else if (
						document
							.querySelector('.gmail_attr')
							?.innerText?.includes('Forwarded message')
					) {
						emailContent = document.querySelector('.gmail_attr')?.innerText;
					}
					if (emailContent?.includes('Forwarded message')) {
						email = emailContent.substring(
							emailContent.indexOf('<') + 1,
							emailContent.indexOf('>'),
						);
					}
					if (
						email?.includes('@ineuron.ai') ||
						email?.includes('@pwskills.com') ||
						email?.includes('@pw.live')
					) {
						email = document
							.querySelectorAll('.gF.gK')[0]
							.firstChild.getAttribute('email');
					}
				}
				let date = new Date(document.getElementsByClassName('g3')[0]?.title);
				let subject = document.getElementsByClassName('hP')[0]?.innerText;
				let smeName = localStorage.getItem('name');
				if (!localStorage.getItem('name')) {
					const name = prompt('Please enter your name');
					localStorage.setItem('name', name);
					smeName = name;
				}
				let fullEmail = document.getElementsByClassName('gb_d')[1]?.ariaLabel;
				if (fullEmail) {
					fullEmail = fullEmail.substring(
						fullEmail.indexOf('(') + 1,
						fullEmail.indexOf(')'),
					);
					if (fullEmail === 'query@ineuron.ai') {
						fullEmail = 'Query_ineuron';
					} else if (fullEmail === 'support@pwskills.com') {
						fullEmail = 'Pw_Skills';
					}
				}
				let platform = fullEmail;
				let mailType = 'Non-technical';
				let mailStatus = 'Open';
				const data = {
					email,
					date,
					platform,
					subject,
					mailType,
					smeName,
					mailStatus,
				};
				respondedBtn.style.width = '130px';
				respondedBtn.style.margin = '10px';
				respondedBtn.style.height = '40px';
				respondedBtn.style.color = '#fff';
				respondedBtn.style.borderRadius = '5px';
				respondedBtn.style.padding = '10px 25px';
				respondedBtn.style.fontFamily = "'Lato', sans-serif";
				respondedBtn.style.fontWeight = '500';
				respondedBtn.style.background = '#030303';
				respondedBtn.style.cursor = 'pointer';
				respondedBtn.style.transition = 'all 0.3s ease';
				respondedBtn.style.position = 'relative';
				respondedBtn.style.display = 'inline-block';
				respondedBtn.style.outline = 'none';
				// Create the select element
				const selectElement = document.createElement('select');
				// Define the options
				const options = [
					'Non-technical',
					'Technical',
					'Feedback',
					'Escalation',
				];
				// Create and append options to the select element
				options.forEach((optionText) => {
					const option = document.createElement('option');
					option.text = optionText;
					selectElement.appendChild(option);
				});
				selectElement.style.zIndex = 1000;
				selectElement.style.padding = '10px';
				selectElement.style.marginRight = '10px';
				selectElement.style.marginLeft = '15px';
				selectElement.addEventListener('change', (event) => {
					data.mailType = event.target.value;
				});
				const emailInput = document.createElement('input');
				emailInput.type = 'text';
				emailInput.value = email;
				emailInput.addEventListener('change', (event) => {
					emailInput.value = event.target.value;
					email = event.target.value;
					data.email = event.target.value;
				});
				emailInput.style.padding = '10px 5px';
				emailInput.style.marginLeft = '10px';
				container?.appendChild(emailInput);
				console.log('Input element created');
				container?.appendChild(selectElement);
				console.log('Select element created');
				container?.appendChild(respondedBtn);
				console.log('Responded button added to container');
				respondedBtn.addEventListener('click', () => {
					postData(data);
					console.log('data posted', data);
				});
			}
			if (mutation.type === 'childList') {
				console.log('Child elements have been added or removed.');
			} else if (mutation.type === 'attributes') {
				console.log('Attribute has changed.');
			}
		}
	});
	// Configuration for the observer
	const config = {
		childList: true, // Observe additions/removals of child elements
		subtree: true, // Observe changes in descendant elements as well
	};
	// Start observing the target element with the provided configuration
	observer.observe(targetElement, config);
});

const form = document.createElement('form');

form.addEventListener('submit', (e) => {
	e.preventDefault();
	// Send the data to the server

	onsubmit();
});
