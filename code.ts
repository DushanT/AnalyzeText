// This file holds the main code for the plugin. It has access to the *document*.
// You can access browser APIs such as the network by creating a UI which contains
// a full browser environment (see documentation).

// Runs this code if the plugin is run in Figma
if (figma.editorType === 'figma') {

  // This shows the HTML page in "ui.html".
  figma.showUI(__html__, {
    themeColors: true,
    title: 'TextFix',
    width: 350,
  });

  // For storing data about library styles
  let styles: { id: string, name: string }[] = []

  // @ts-ignore
  const isTextNode = node => node.type === 'TEXT'
  // @ts-ignore
  const isTextWithEmptyStyleId = node => isTextNode(node) && node.textStyleId === ''
  // @ts-ignore
  const isTextWithStyleId = node => isTextNode(node) && node.textStyleId

  // Initialize data from client storage
  const init = async () => {
    // if (figma.hasMissingFont) {
    //   figma.ui.postMessage({
    //     type: 'missing-fonts'
    //   })
    // }
    styles = JSON.parse(await figma.clientStorage.getAsync('styles'))
    if (styles.length > 0) {
      figma.ui.postMessage({
        type: 'loaded-styles',
        styles
      })
      try {
        await Promise.all(
          // @ts-ignore
          styles.map(style => figma.loadFontAsync(style.fontName))
        )
      } catch(err) {
        console.error(`Error loading fonts: ${err}`);
        figma.notify(`Error loading fonts: ${err}`, { error: true })
      }
      figma.ui.resize(350, 450)
    }
  }

  // Save style data to client storage
  const loadStyles = () => {
    // find all unique nested texts with style
    const textNodes = figma.currentPage.selection
      // @ts-ignore
      .reduce((acc, node) => {
        // @ts-ignore
        return [...acc, ...(node.children ? node.findAll(isTextWithStyleId) : [node])]
      }, [])
      // @ts-ignore
      .filter((node, index, array) => {
        // @ts-ignore
        return array.findIndex(n => n.textStyleId === node.textStyleId) === index && isTextWithStyleId(node)
      })
    if (textNodes.length > 0) {
      // @ts-ignore
      const styles = textNodes.map((node) => ({
        // @ts-ignore
        id: node.textStyleId,
        // @ts-ignore
        name: figma.getStyleById(node.textStyleId)?.name,
        fontName: node.fontName
      }))
      figma.clientStorage.setAsync('styles', JSON.stringify(styles))
      figma.ui.postMessage({
        type: 'loaded-styles',
        styles
      })
      figma.ui.resize(350, 450)
      figma.notify(`Done loading! Unique styles loaded: ${textNodes.length}`)
    } else {
      figma.notify('You have to select text nodes with styles', {
        error: true
      })
    }
  }

  // Delete styles client storage data
  const clearStyles = () => {
    figma.clientStorage.deleteAsync('styles')
    figma.ui.postMessage({
      type: 'clear-styles',
    })
    figma.ui.resize(350, 200)
  }

  // Check if style matches node styles
  // @ts-ignore
  const checkEquality = (style, node) => {
    return (
      style && node &&
      style.fontSize === node.fontSize &&
      style.letterSpacing.unit === node.letterSpacing.unit &&
      style.letterSpacing.value === node.letterSpacing.value &&
      style.lineHeight.unit === node.lineHeight.unit &&
      style.lineHeight.value === node.lineHeight.value &&
      style.listSpacing === node.listSpacing &&
      style.paragraphIndent === node.paragraphIndent &&
      style.paragraphSpacing === node.paragraphSpacing &&
      style.textCase === node.textCase &&
      // // text decoration can be broken cause of mixed styles applied to text
      // // for example text with some part manually underlined
      // !!! Changing text segments always detach text style from text, this is useless
      // !!! We should leave the manually adjusted text detached or not manually 
      // !!! change underlines and boldes inside text nodes
      // (
      //   style.textDecoration === node.textDecoration || (
      //     typeof node.textDecoration !== 'string' &&
      //     node.getStyledTextSegments(['textDecoration'])?.length > 1
      //   )
      // ) &&
      // (
      //   (
      //     style.fontName.family === node.fontName.family &&
      //     style.fontName.style === node.fontName.style    
      //   ) || 
      //   (
      //     typeof node.fontName !== 'string' &&
      //     node.getStyledTextSegments(['fontName'])?.length > 1
      //   )
      // )
      /** 
       * if someone changes part of the text to underline it's not possible 
       * to attach a text style to this node because text properties are mixed
       */ 
      style.textDecoration === node.textDecoration && 
      style.fontName.family === node.fontName.family &&
      style.fontName.style === node.fontName.style  
    )
  }

  // Find and select all nodes with equal text styles by style id 
  const findAndSelectNodes = (id: string, wrapperNode = figma.currentPage, zoom = false) => {
    if (!wrapperNode) {
      figma.notify(`No wrapper defined`, { error: true })
      return
    }

    const style = figma.getStyleById(id)
    if (!style) {
      figma.notify(`No style found for id: ${id}`, { error: true })
      return
    }

    const nodes = wrapperNode.findAll(node => isTextWithEmptyStyleId(node))
    const matchingNodes = nodes.filter(node => checkEquality(style, node))

    figma.currentPage.selection = matchingNodes
    if (matchingNodes.length > 0) {
      if (zoom) {
        figma.viewport.scrollAndZoomIntoView(figma.currentPage.selection)
      }
      figma.notify(`Detached nodes: ${matchingNodes.length}`)
    } else {
      figma.notify('No detached nodes found', { timeout: 1000, error: true })
    }
  }

  // Find matching style from array of styles
  const findMatchingStyle = (node: DocumentNode | PageNode) => {
    return styles.find(styleObj => {
      const style = figma.getStyleById(styleObj.id);
      return checkEquality(style, node)
    })?.id
  }

  // Find matching nodes on the current page
  const findCurrentNodes = (id: string) => {
    findAndSelectNodes(id, figma.currentPage, true)
  }

  // Apply styles to node by style id
  const applyStyle = (id: string, node: TextNode) => {
    // const textSegments = node.getStyledTextSegments(['listOptions', 'textCase', 'textDecoration', 'fontName'])
    node.textStyleId = id
    // !!! Changing text segments always detach text style from text, this is useless 
    // !!! because if we change text segments it will detach the style again
    // if (textSegments.length > 1) {
    //   textSegments.map(segment => {
    //     node.setRangeListOptions(segment.start, segment.end, segment.listOptions)
    //     node.setRangeTextDecoration(segment.start, segment.end, segment.textDecoration)
    //     node.setRangeTextCase(segment.start, segment.end, segment.textCase)
    //     node.setRangeFontName(segment.start, segment.end, segment.fontName)
    //   })
    // }
  }

  // Apply style to current page selection
  const applyStyleToSelection = (id: string) => {
    const processedCount = figma.currentPage.selection.length
    if (processedCount > 0) {
      // @ts-ignore
      figma.currentPage.selection.forEach(node => applyStyle(id, node))
      figma.notify(`Done applying styles! Processed nodes: ${processedCount}`)
    } else {
      figma.notify('You have to select/find some nodes to apply styles', { error: true })
    }
  }

  // Find nodes without styles
  const findNodesWithoutStyles = () => {
    const noStyleTexts = figma.currentPage.findAll(node => isTextWithEmptyStyleId(node))
    figma.currentPage.selection = noStyleTexts
    figma.viewport.scrollAndZoomIntoView(figma.currentPage.selection)
    figma.notify(`Done searching! Selected nodes: ${noStyleTexts.length}`)
  }

  const findMatchingStyleNodes = (styleId: string) => {
    const style = figma.getStyleById(styleId)
    const nodes = figma.currentPage
      .findAll(node => isTextWithEmptyStyleId(node))
      .filter(node => checkEquality(style, node))
    figma.currentPage.selection = nodes
    figma.viewport.scrollAndZoomIntoView(figma.currentPage.selection)
    figma.notify(`Done searching! Selected nodes: ${nodes.length}`)
  }

  // Find node style and apply it
  // @ts-ignore
  const findAndApplyStyle = (node) => {
    const styleId = findMatchingStyle(node)
    if (styleId) {
      applyStyle(styleId, node)
      return true
    }
    return false
  }

  // fix detached nodes
  const fixDetached = (wrapperNode = figma.currentPage) => {
    let countSuccessful = 0
    const nodes = wrapperNode
      .findAll(node => isTextWithEmptyStyleId(node))
    nodes
      .forEach(node => {
        if (findAndApplyStyle(node)) {
          countSuccessful += 1
        }
      })
    figma.notify(`Done replacing! Processed: ${nodes.length}, Replaced: ${countSuccessful}`)
  }

  // fix all pages detached nodes
  const fixAllDetached = () => {
    // @ts-ignore
    fixDetached(figma.root)
  }

  // fix current page detached nodes
  const fixCurrentDetached = () => {
    fixDetached(figma.currentPage)
  }

  // fix current page detached nodes
  const fixSelected = () => {
    let countSuccessful = 0
    figma.currentPage.selection.forEach(node => {
      if (isTextWithEmptyStyleId(node) && findAndApplyStyle(node)) {
        countSuccessful += 1
      }
    })
    const countProcessed = figma.currentPage.selection.length
    if (countProcessed > 0) {
      figma.notify(`Done replacing! Nodes processed: ${countProcessed}, Nodes fixed: ${countSuccessful}`)
    } else {
      figma.notify('You have to select something to apply styles', { error: true })
    }
  }

  let zoomTargetId = 0

  // Animate fill of text
  // @ts-ignore
  const animateFill = (target) => {
    const currentFills = target.fills
    target.fills = [{ type: 'SOLID', color: { r: 1, g: 0, b: 0 } }]
    setTimeout(() => {
      target.fills = currentFills
      setTimeout(() => {
        target.fills = [{ type: 'SOLID', color: { r: 1, g: 0, b: 0 } }]
        setTimeout(() => {
          target.fills = currentFills
        }, 300)
      }, 150)
    }, 300)
  }

  // Zoom to node
  const zoomTo = (increment: number) => {
    const targetId = zoomTargetId + increment
    const sel = figma.currentPage.selection
    zoomTargetId = targetId < 0 ? (sel.length + targetId) : targetId % sel.length
    const target = sel[zoomTargetId]
    if (target) {
      figma.viewport.scrollAndZoomIntoView([target])
      animateFill(target)
    } else {
      figma.notify('Noting to zoom in, select something', { error: true })
    }
  }

  init()

  figma.on('currentpagechange', () => {
    zoomTargetId = 0
  })

  figma.on('selectionchange', () => {
    const selectionLength = figma.currentPage.selection.length
    figma.ui.postMessage({
      type: 'selection-changed',
      selectionLength
    })
  })

  // Calls to "parent.postMessage" from within the HTML page will trigger this
  // callback. The callback will be passed the "pluginMessage" property of the
  // posted message.
  figma.ui.onmessage = msg => {
    // One way of distinguishing between different types of messages sent from
    // your HTML page is to use an object with a "type" property like this.

    switch (msg.type) {
      case 'load-base-styles':
        loadStyles();
        break;
      case 'clear-base-styles':
        clearStyles();
        break;
      case 'find-current-nodes':
        findCurrentNodes(msg.id);
        break;
      case 'apply-style':
        applyStyleToSelection(msg.id);
        break;
      case 'find-no-style-nodes':
        findNodesWithoutStyles();
        break;
      case 'find-matching-style-nodes':
        findMatchingStyleNodes(msg.id);
        break;
      case 'zoom-prev':
        zoomTo(-1);
        break;
      case 'zoom-next':
        zoomTo(1);
        break;
      case 'fix-selected':
        fixSelected()
        break;
      case 'fix-all-detached':
        fixAllDetached();
        break;
      case 'fix-current-detached':
        fixCurrentDetached();
        break;
      case 'close':
        // Make sure to close the plugin when you're done. Otherwise the plugin will
        // keep running, which shows the cancel button at the bottom of the screen.
        figma.closePlugin();
        break;
      default:
        // handle default case
        break;
    }
  };
};
