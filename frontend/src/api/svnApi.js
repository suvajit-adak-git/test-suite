const API_URL = 'http://localhost:8000/api'; // Assuming the backend is running on port 8000

export const uploadSvnFile = async (file) => {
    const formData = new FormData();
    formData.append('file', file);

    const response = await fetch(`${API_URL}/upload-excel`, {
        method: 'POST',
        body: formData,
    });

    if (!response.ok) {
        throw new Error('Failed to upload SVN file');
    }

    return response.json();
};

export const uploadReviewChecklist = async (file) => {
    const formData = new FormData();
    formData.append('file', file);

    const response = await fetch(`${API_URL}/upload-review-checklist`, {
        method: 'POST',
        body: formData,
    });

    if (!response.ok) {
        throw new Error('Failed to upload review checklist');
    }

    return response.json();
};

export const compareBoth = async (payload) => {
    const response = await fetch(`${API_URL}/compare-both`, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify(payload),
    });

    if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.detail || 'Failed to compare files');
    }

    return response.json();
};
export const extractHyperlinks = async (file) => {
    const formData = new FormData();
    formData.append('file', file);

    const response = await fetch(`${API_URL}/hyperlinks/extract-hyperlinks/`, {
        method: 'POST',
        body: formData,
    });

    if (!response.ok) {
        throw new Error('Failed to extract hyperlinks');
    }

    return response.json();
};

export const updateBuild = async (file, newBuild) => {
    const formData = new FormData();
    formData.append('file', file);
    formData.append('new_build', newBuild);

    const response = await fetch(`${API_URL}/hyperlinks/update-build/`, {
        method: 'POST',
        body: formData,
    });

    if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.detail || 'Failed to update build');
    }

    return response.blob();
};
