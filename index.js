require('dotenv').config()
const { default: axios } = require('axios')
const express = require('express')
const app = express()

const router = express.Router()

app.use(router)
app.use(express.json())

app.get('/', (req, res, next) => {
  try {
    res.send('hello world')
  } catch (error) {
    res.status(500).send(error)
    next(error)
  }
})

app.post('/simulateError', (req, res, next) => {
  try {
    throw {
      status: 400,
      msg: 'Something went wrong',
    }
  } catch (error) {
    res.status(error.status).send(error)

    // calling next here , control flows to next middleware in the chain
    // which is the errorNotifier
    next(error)
  }
})

const errorNotifier = async (err, req, res, next) => {
  let notifierCardTemplate = {
    '@type': 'MessageCard',
    '@context': 'http://schema.org/extensions',
    themeColor: '0076D7',
    summary: 'Application Errors',
    sections: [
      {
        activityTitle: 'Application Errors',
        activitySubtitle: 'Service',
        activityImage:
          'https://teamsnodesample.azurewebsites.net/static/img/image5.png',
        facts: [
          {
            name: 'Path',
            value: req.path,
          },
          {
            name: 'Request data',
            value: ` \`\`\`${JSON.stringify(req.body)}  \`\`\` `,
          },
          {
            name: 'Error Message',
            value: ` \`\`\`${JSON.stringify(err)}  \`\`\` `,
          },
        ],
        markdown: true,
      },
    ],
    potentialAction: [
      {
        '@type': 'ActionCard',
        name: 'Add a comment',
        inputs: [
          {
            '@type': 'TextInput',
            id: 'comment',
            isMultiline: false,
            title: 'Add a comment here for this task',
          },
        ],
        actions: [
          {
            '@type': 'HttpPOST',
            name: 'Add comment',
            target: 'https://docs.microsoft.com/outlook/actionable-messages',
          },
        ],
      },
      {
        '@type': 'ActionCard',
        name: 'Set due date',
        inputs: [
          {
            '@type': 'DateInput',
            id: 'dueDate',
            title: 'Enter a due date for this task',
          },
        ],
        actions: [
          {
            '@type': 'HttpPOST',
            name: 'Save',
            target: 'https://docs.microsoft.com/outlook/actionable-messages',
          },
        ],
      },
      {
        '@type': 'OpenUri',
        name: 'Learn More',
        targets: [
          {
            os: 'default',
            uri: 'https://docs.microsoft.com/outlook/actionable-messages',
          },
        ],
      },
      {
        '@type': 'ActionCard',
        name: 'Change status',
        inputs: [
          {
            '@type': 'MultichoiceInput',
            id: 'list',
            title: 'Select a status',
            isMultiSelect: 'false',
            choices: [
              {
                display: 'In Progress',
                value: '1',
              },
              {
                display: 'Active',
                value: '2',
              },
              {
                display: 'Closed',
                value: '3',
              },
            ],
          },
        ],
        actions: [
          {
            '@type': 'HttpPOST',
            name: 'Save',
            target: 'https://docs.microsoft.com/outlook/actionable-messages',
          },
        ],
      },
    ],
  }

  const options = {
    method: 'post',
    url: process.env.WEB_HOOK_URL,
    data: notifierCardTemplate,
  }

  try {
    await axios(options)
  } catch (error) {
    console.log('Error while sending notification')
  }
}

// last middleware in the chain
app.use(errorNotifier)

app.listen(process.env.PORT, () => [
  console.log('Server initialized at port', process.env.PORT),
])
