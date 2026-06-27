const form = document.querySelector('#share-form');

if (form) {
  form.addEventListener('submit', (event) => {
    event.preventDefault();
    const name = form.name.value.trim();
    const title = form.title.value.trim();
    const story = form.story.value.trim();

    if (!name || !title || !story) {
      alert('Please complete every field before submitting.');
      return;
    }

    alert(`Thank you, ${name}! Your piece titled "${title}" has been captured as a local draft. You can integrate a backend or email service to store submissions later.`);
    form.reset();
  });
}
