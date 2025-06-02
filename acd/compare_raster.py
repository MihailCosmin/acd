import skimage.metrics
import skimage.transform

def calculate_image_similarity(image1_path, image2_path):
    # Load images
    try:
        image1 = skimage.io.imread(image1_path, as_gray=True)
    except SyntaxError:  # Cosmin: added this because one emf file was not correctly read
        return 0
    try:
        image2 = skimage.io.imread(image2_path, as_gray=True)
    except SyntaxError:  # Cosmin: added this because one emf file was not correctly read
        return 0

    # print(f"image1.shape: {image1.shape}")
    # print(f"image2.shape: {image2.shape}")

    # find which is the smallest image and resize the other one
    if image1.shape[0] * image1.shape[1] > image2.shape[0] * image2.shape[1]:
        image1 = skimage.transform.resize(image1, image2.shape)
    else:
        image2 = skimage.transform.resize(image2, image1.shape)

    # Calculate SSIM score
    ssim_score = skimage.metrics.structural_similarity(
        image1, image2, data_range=image1.max() - image1.min())
    return ssim_score * 100

if __name__ == "__main__":
    print(calculate_image_similarity(
        r"C:\Users\munteanu\Downloads\_CGM_TIFF\31-17-12\A320A321_IPC_311712_097Z_1_03_from_raster.jpg",
        # r"C:\Users\munteanu\Downloads\_CGM_TIFF\31-17-12\A320A321_IPC_311712_097Z_1_03.tif",
        r"C:\Users\munteanu\Downloads\_CGM_TIFF\31-17-12\A320A321_IPC_311712_097Z_1_03_from_cgm.jpg"
    ))