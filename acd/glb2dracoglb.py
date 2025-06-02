"""This module provides functions to perform operations on glb/gltf files
"""
from os import system
from os import remove
from os.path import join

from tqdm import tqdm

from .filelist import list_files  #: :noindex: 

def glb2dracoglb(directory: str, debug: bool = False) -> None:
    """This function converts the glb files from a directory to draco glb files (compressed).
    
    Args:
        directory (str): The directory to convert the glb files from.
        debug (bool, optional): If True, the function will print debug messages. Defaults to False.
        
    Returns:
        None - it performs the conversion in place.
    """
    for file_path in tqdm(list_files(directory, True)):
        if file_path.endswith(".glb"):
            glb = join(directory, file_path)
            gltf = join(directory, file_path.replace('.glb', '.gltf'))
            if debug:
                print(f'Converting "{glb}" to "{gltf}"')
            system(f'gltf-pipeline -i "{glb}" -o "{gltf}" --draco.compressionLevel')
            if debug:
                print(f'Removing "{glb}"')
            remove(glb)
            if debug:
                print(f'Converting "{gltf}" to "{glb}" (draco compressed glb)')
            system(f'gltf-pipeline -i "{gltf}" -o "{glb}" --binary --draco')
            if debug:
                print(f'Removing "{gltf}"')
            remove(gltf)
            if debug:
                print(f'Finished converting "{glb}" to draco compressed glb')

    print('Finished converting all glb files to draco compressed glb files.')

if __name__ == "__main__":
    glb2dracoglb(r"C:\Users\munteanu\Desktop\Three.js  Configurator\Draco\GLB to DRACO GLB")
